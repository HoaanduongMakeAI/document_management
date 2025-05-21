frappe.ui.form.on('Outgoing Document', {
    custom_handle_attach_and_upload: function(frm) {
        let current_sharepoint_path = '/'; // Relative path within the selected_folder, always starts at root of selected_folder
        let current_selected_folder_docname = frm.doc.folder || ''; // Base Folder Doctype name
        let selected_item_for_linking = null; // Stores { name: '...', path: '...', is_folder: false }

        const dialog = new frappe.ui.Dialog({
            title: __('Upload or Link File to SharePoint'),
            fields: [
                {
                    fieldname: 'selected_folder',
                    fieldtype: 'Link',
                    label: __('SharePoint Folder Location'),
                    options: 'Folder',
                    default: frm.doc.folder,
                    reqd: 1,
                    description: __('Select the base SharePoint Folder record. This determines the root for browsing.'),
                    onchange: function() {
                        let new_folder_docname = this.get_value();
                        if (new_folder_docname) {
                            current_selected_folder_docname = new_folder_docname;
                            current_sharepoint_path = '/';
                            selected_item_for_linking = null;
                            dialog.get_field('file_to_upload').toggle(true);
                            dialog.get_field('unlink_button').toggle(false);
                            update_path_display();
                        } else {
                            current_selected_folder_docname = '';
                            dialog.get_field('path_display_html').$wrapper.html('');
                            dialog.get_field('back_button').toggle(false);
                        }
                    }
                },
                {
                    fieldname: 'path_display_html', // Will display list and selection info
                    fieldtype: 'HTML',
                    label: __('Current Path Contents')
                },
                {
                    fieldname: 'back_button',
                    fieldtype: 'Button',
                    label: __('Back to Parent Folder'),
                    hidden: 1,
                    click: function() {
                        if (current_sharepoint_path && current_sharepoint_path !== '/') {
                            let parts = current_sharepoint_path.split('/').filter(p => p.trim() !== '');
                            parts.pop(); // Go up one level
                            current_sharepoint_path = '/' + parts.join('/');
                            if (current_sharepoint_path === '//') current_sharepoint_path = '/';
                            selected_item_for_linking = null;
                            dialog.get_field('file_to_upload').toggle(true);
                            dialog.get_field('unlink_button').toggle(false);
                            update_path_display(); // This will clear the selection message by re-rendering
                        }
                    }
                },
                {
                    fieldname: 'unlink_button',
                    fieldtype: 'Button',
                    label: __('Clear Selection / Upload'),
                    hidden: 1,
                    click: function() {
                        selected_item_for_linking = null;
                        dialog.set_value('file_to_upload', null);
                        
                        let attach_field = dialog.get_field('file_to_upload');
                        if (attach_field && attach_field.uploader) {
                            attach_field.uploader.reset();
                        }

                        dialog.get_field('file_to_upload').toggle(true);
                        dialog.get_field('unlink_button').toggle(false); // Correctly hide the button
                        update_path_display();
                    }
                },
                {
                    fieldname: 'file_to_upload',
                    fieldtype: 'Attach',
                    label: __('Upload New File (Optional)'),
                    description: __('If you want to upload a new file to the selected path/folder.'),
                    onchange: function() {
                        // When a file is selected or cleared in the attach field, update the display
                        // This will ensure the "unlink" button visibility and messages are correct.
                        update_path_display();
                    }
                }
            ],
            primary_action_label: __('Process Selection'),
            primary_action: function(values) {
                let file_to_upload_url = dialog.get_value('file_to_upload');
                let target_folder_docname_for_action = dialog.get_value('selected_folder');

                if (!target_folder_docname_for_action) {
                    frappe.msgprint(__('Please select a SharePoint Folder Location.'));
                    return;
                }

                // Scenario 1: Uploading a new file
                if (file_to_upload_url) {
                    frappe.show_alert({ message: __('Fetching file details for upload...'), indicator: 'info' });
                    frappe.call({
                        method: 'document_management.document_management.utils.sharepoint_integration.get_latest_file_doc_name_by_url',
                        args: { file_url: file_to_upload_url },
                        callback: function(r_file_doc) {
                            if (r_file_doc.message) {
                                let file_doc_name = r_file_doc.message;
                                frappe.show_alert({ message: __('File details fetched. Uploading to SharePoint...'), indicator: 'info' });
                                // current_sharepoint_path here should be the folder path to upload into.
                                // If a file was selected for linking, current_sharepoint_path points to that file.
                                // We need the parent folder of that file, or the current browsed path if it's a folder.
                                let upload_target_relative_path = current_sharepoint_path;
                                if (selected_item_for_linking && !selected_item_for_linking.is_folder) {
                                    // If a file was selected, upload to its parent directory
                                    let parts = selected_item_for_linking.path.split('/').filter(p => p.trim() !== '');
                                    parts.pop();
                                    upload_target_relative_path = '/' + parts.join('/');
                                    if (upload_target_relative_path === '//') upload_target_relative_path = '/';
                                }


                                frappe.call({
                                    method: 'document_management.document_management.utils.sharepoint_integration.upload_file_to_path',
                                    args: {
                                        doctype: frm.doc.doctype,
                                        docname: frm.doc.name,
                                        file_doc_name: file_doc_name,
                                        target_folder_docname: target_folder_docname_for_action,
                                        target_relative_path: upload_target_relative_path
                                    },
                                    callback: function(r_upload) {
                                        if (r_upload.message && r_upload.message.sharepoint_link && r_upload.message.absolute_path) {
                                            frm.set_value('teams_link', r_upload.message.sharepoint_link);
                                            frm.set_value('path', r_upload.message.absolute_path);
                                            frm.set_value('folder', target_folder_docname_for_action);
                                            frappe.show_alert({ message: __('File successfully uploaded and linked. Creating document version...'), indicator: 'info' });
                                            
                                        } else if (r_upload.message && r_upload.message.error) {
                                            frappe.msgprint({ title: __('Upload Error'), indicator: 'red', message: r_upload.message.error });
                                        } else if (r_upload.exc) {
                                            frappe.msgprint({ title: __('Server Error'), indicator: 'red', message: __('An error occurred during SharePoint upload. Check server logs.')});
                                            console.error("SharePoint Upload Error:", r_upload.exc);
                                        } else {
                                            frappe.msgprint({ title: __('Upload Issue'), indicator: 'orange', message: __('Upload completed but no link/path was returned.')});
                                        }
                                        dialog.hide();
                                    },
                                    error: function(err_upload) {
                                        frappe.msgprint({ title: __('Network Error'), indicator: 'red', message: __('Failed to communicate for SharePoint upload.')});
                                        console.error("AJAX Error (Upload):", err_upload);
                                        dialog.hide();
                                    }
                                });
                            } else if (r_file_doc.exc) {
                                frappe.msgprint({ title: __('Server Error'), indicator: 'red', message: __('Error fetching file details. Check server logs.')});
                                console.error("Get File Details Error:", r_file_doc.exc);
                            } else {
                                frappe.msgprint({ title: __('File Details Issue'), indicator: 'orange', message: __('Could not retrieve file details for upload.')});
                            }
                        },
                        error: function(err_file_doc) {
                            frappe.msgprint({ title: __('Network Error'), indicator: 'red', message: __('Failed to communicate for file details.')});
                            console.error("AJAX Error (File Details):", err_file_doc);
                        }
                    });
                }
                // Scenario 2: Linking to an existing file selected via selected_item_for_linking
                else if (selected_item_for_linking && !selected_item_for_linking.is_folder) {
                    frappe.show_alert({ message: __('Fetching details for selected SharePoint item...'), indicator: 'info' });
                    frappe.call({
                        method: 'document_management.document_management.utils.sharepoint_integration.get_sharepoint_item_details',
                        args: {
                            target_folder_docname: target_folder_docname_for_action,
                            relative_path_to_item: selected_item_for_linking.path // Path is relative to Folder DocType root
                        },
                        callback: function(r_item) {
                            if (r_item.message && r_item.message.is_file && r_item.message.sharepoint_link && r_item.message.absolute_path) {
                                frm.set_value('teams_link', r_item.message.sharepoint_link);
                                frm.set_value('path', r_item.message.absolute_path);
                                frm.set_value('folder', target_folder_docname_for_action);
                                frappe.show_alert({ message: __('Successfully linked to existing SharePoint file. Please save the document.'), indicator: 'green' });
                                // User should save manually after mandatory fields are filled.
                                dialog.hide();
                            } else if (r_item.message && !r_item.message.is_file) { // Should not happen if selected_item_for_linking.is_folder is false
                                frappe.msgprint({ title: __('Selection Error'), indicator: 'orange', message: __('The selected item is unexpectedly a folder. Please try again.')});
                            } else if (r_item.message && r_item.message.error) {
                                frappe.msgprint({ title: __('Item Details Error'), indicator: 'red', message: r_item.message.error });
                            } else if (r_item.exc) {
                                frappe.msgprint({ title: __('Server Error'), indicator: 'red', message: __('Error fetching item details. Check server logs.')});
                                console.error("Get Item Details Error:", r_item.exc);
                            } else {
                                frappe.msgprint({ title: __('Item Details Issue'), indicator: 'orange', message: __('Could not retrieve details for the selected SharePoint item.')});
                            }
                        },
                        error: function(err_item) {
                            frappe.msgprint({ title: __('Network Error'), indicator: 'red', message: __('Failed to communicate for item details.')});
                            console.error("AJAX Error (Item Details):", err_item);
                        }
                    });
                } else {
                    frappe.msgprint(__('Please either select a file to upload, or navigate to and select an existing file in SharePoint to link.'));
                }
            }
        });

        function update_path_display() {
            if (!current_selected_folder_docname) {
                dialog.get_field('path_display_html').$wrapper.html('<div class="text-muted">' + __('Please select a SharePoint Folder Location first.') + '</div>');
                dialog.get_field('back_button').toggle(false);
                return;
            }
            let loading_path_display = current_sharepoint_path === '/' ? current_selected_folder_docname : `${current_selected_folder_docname}${current_sharepoint_path}`;
            dialog.get_field('path_display_html').$wrapper.html(`<div class="text-muted"><i class="fa fa-spinner fa-spin"></i> ${__('Loading contents for:')} <strong>${loading_path_display}</strong></div>`);

            frappe.call({
                method: "document_management.document_management.utils.sharepoint_integration.list_sharepoint_folder_contents",
                args: {
                    folder_docname: current_selected_folder_docname,
                    relative_path: current_sharepoint_path // This is relative to the Folder DocType's root
                },
                callback: function(r) {
                    let file_is_staged_for_upload = dialog.get_value('file_to_upload');
                    let file_is_selected_for_linking = selected_item_for_linking && !selected_item_for_linking.is_folder;

                    dialog.get_field('unlink_button').toggle(file_is_selected_for_linking || file_is_staged_for_upload);

                    if (file_is_selected_for_linking) {
                        dialog.get_field('file_to_upload').toggle(false);
                    } else {
                        dialog.get_field('file_to_upload').toggle(true);
                    }

                    let html = `<p class="text-muted small mb-2">${__('Current Location:')} <strong>${loading_path_display}</strong></p>`;
                    if (r.message && r.message.items) {
                        if (r.message.items.length === 0) {
                            html += '<div class="text-muted">' + __('This folder is empty.') + '</div>';
                        } else {
                            html += '<ul class="list-group">';
                            r.message.items.forEach(item => {
                                let item_class = "sharepoint-item";
                                let item_style = "text-decoration: none; color: inherit; display: block;";
                                if (selected_item_for_linking && selected_item_for_linking.path === item.path && !selected_item_for_linking.is_folder) {
                                    item_class += " active";
                                }
                                html += `<li class="list-group-item list-group-item-action p-2">
                                            <a href="#" class="${item_class}" data-path="${item.path}" data-name="${item.name}" data-is-folder="${item.is_folder}" style="${item_style}">
                                                <i class="fa ${item.is_folder ? 'fa-folder' : 'fa-file-o'} mr-2"></i> ${item.name}
                                            </a>
                                         </li>`;
                            });
                            html += '</ul>';
                        }

                        // Display messages based on current state
                        if (file_is_staged_for_upload) {
                            let upload_target_display_path = current_sharepoint_path === '/' ? current_selected_folder_docname : `${current_selected_folder_docname}${current_sharepoint_path}`;
                            html += `<p class="text-success small mt-2 p-2"><strong>${__("New file will be uploaded to:")} ${upload_target_display_path}</strong></p>`;
                        } else if (file_is_selected_for_linking) {
                            html += `<p class="text-info small mt-2 p-2"><strong>${__("Selected for linking:")} ${selected_item_for_linking.name}</strong> (${selected_item_for_linking.path})</p>`;
                        } else if (current_selected_folder_docname && r.message && typeof r.message.items !== 'undefined') {
                            // Only show this if a folder is loaded (items array exists) and nothing else is selected/staged
                            html += `<p class="text-warning small mt-2 p-2">${__("Please upload a new file or select an existing file from the list to link.")}</p>`;
                        }
                        
                        dialog.get_field('path_display_html').$wrapper.html(html);
                        dialog.get_field('back_button').toggle(current_sharepoint_path !== '/' && current_sharepoint_path !== '');

                        dialog.get_field('path_display_html').$wrapper.find('a.sharepoint-item').on('click', function(e) {
                            e.preventDefault();
                            let new_relative_path_in_folder = $(this).data('path');
                            let is_folder = $(this).data('is-folder');
                            let item_name = $(this).data('name');

                            if (is_folder) {
                                current_sharepoint_path = new_relative_path_in_folder;
                                selected_item_for_linking = null;
                                // Do not clear file_to_upload here, user might want to upload to this new folder.
                                // update_path_display will handle visibility of unlink_button and file_to_upload field.
                                update_path_display();
                            } else {
                                // File selected for linking
                                selected_item_for_linking = { name: item_name, path: new_relative_path_in_folder, is_folder: false };
                                dialog.set_value('file_to_upload', null); // Clear any staged upload
                                let attach_field_sel = dialog.get_field('file_to_upload');
                                if (attach_field_sel && attach_field_sel.uploader) {
                                    attach_field_sel.uploader.reset();
                                }
                                // update_path_display will hide the upload field and show unlink button.
                                update_path_display();
                                frappe.show_alert({
                                    message: __("File '{0}' is selected for linking. If 'Process Selection' is clicked now, this file will be linked. To upload a new file instead, clear the selection first.", [item_name]),
                                    indicator: 'info',
                                    toast: true,
                                    display_length: 10000
                                });
                            }
                        });
                    } else if (r.message && r.message.error) {
                        dialog.get_field('path_display_html').$wrapper.html(html + `<div class="text-danger mt-2">${__('Error')}: ${r.message.error}</div>`);
                        dialog.get_field('back_button').toggle(current_sharepoint_path !== '/' && current_sharepoint_path !== '');
                    } else if (r.exc) {
                        dialog.get_field('path_display_html').$wrapper.html(html + `<div class="text-danger mt-2">${__('Server error while listing contents.')}</div>`);
                        console.error("List Contents Error:", r.exc);
                        dialog.get_field('back_button').toggle(current_sharepoint_path !== '/' && current_sharepoint_path !== '');
                    } else {
                        dialog.get_field('path_display_html').$wrapper.html(html + '<div class="text-muted mt-2">' + __('Folder is empty or unable to load contents.') + '</div>');
                        dialog.get_field('back_button').toggle(current_sharepoint_path !== '/' && current_sharepoint_path !== '');
                    }
                },
                error: function(err) {
                    dialog.get_field('path_display_html').$wrapper.html(html + `<div class="text-danger mt-2">${__('Network error while listing contents.')}</div>`);
                    console.error("AJAX Error (List Contents):", err);
                    dialog.get_field('back_button').toggle(current_sharepoint_path !== '/' && current_sharepoint_path !== '');
                }
            });
        }

        dialog.show();
        if (current_selected_folder_docname) {
            update_path_display();
        } else {
            dialog.get_field('path_display_html').$wrapper.html('<div class="text-muted">' + __('Please select a SharePoint Folder Location to browse.') + '</div>');
        }
    },

    refresh: function(frm) {
        // Control visibility of the open_teams_link_btn
        frm.get_field('open_teams_link_btn').toggle(Boolean(frm.doc.teams_link));

        if (!frm.is_new()) {
            frm.add_custom_button(__('Attach/Link SharePoint File'), function() { // Changed button label
                frm.trigger('custom_handle_attach_and_upload');
            }, __('Actions'));
        }
    },

    teams_link: function(frm) {
        // Control visibility of the open_teams_link_btn when teams_link changes
        frm.get_field('open_teams_link_btn').toggle(Boolean(frm.doc.teams_link));
    },

    open_teams_link_btn: function(frm) {
        if (frm.doc.teams_link) {
            window.open(frm.doc.teams_link, '_blank');
        }
    },

    upload_to_sharepoint_btn: function(frm) { // This button is likely from JSON, ensure its label is also updated or it's removed if custom_button is preferred
        if (!frm.is_new()) {
            frm.trigger('custom_handle_attach_and_upload');
        } else {
            frappe.msgprint(__('Please save the document before attaching files.'));
        }
    }
});