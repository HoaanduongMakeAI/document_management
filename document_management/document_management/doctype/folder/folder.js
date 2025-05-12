frappe.ui.form.on('Folder', {
    refresh: function(frm) {
        // Initialize current_sharepoint_path relative to the Folder doctype's path
        let current_sharepoint_path = '/'; // Path relative to frm.doc.folder_path
        let current_folder_docname = frm.doc.name; // The Folder doctype's name is the base

        // Add a "Back" button above the HTML field
        if (!frm.custom_back_button_added) {
            frm.add_custom_button(__('Back to Parent Folder'), function() {
                if (current_sharepoint_path && current_sharepoint_path !== '/') {
                    let parts = current_sharepoint_path.split('/').filter(p => p.trim() !== '');
                    parts.pop(); // Go up one level
                    current_sharepoint_path = '/' + parts.join('/');
                    if (current_sharepoint_path === '//') current_sharepoint_path = '/'; // Normalize if it became just '//'
                    update_sharepoint_contents_display();
                }
            }, __('SharePoint Contents'));
             frm.custom_back_button_added = true; // Prevent adding multiple buttons
        }


        // Function to update the SharePoint contents display
        function update_sharepoint_contents_display() {
            if (!current_folder_docname) {
                 if (frm.fields_dict['sharepoint_contents_html']) {
                     frm.fields_dict['sharepoint_contents_html'].$wrapper.html('<div class="text-muted">' + __('Please save the Folder document first.') + '</div>');
                 }
                 // Check if the field exists before toggling
                 if (frm.get_field('sharepoint_contents_tab')) {
                     frm.get_field('sharepoint_contents_tab').toggle(false); // Hide tab if not saved
                 }
                 return;
            }

            // Check if the field exists before toggling
            if (frm.get_field('sharepoint_contents_tab')) {
                frm.get_field('sharepoint_contents_tab').toggle(true); // Ensure tab is visible
            }

            let base_dp = frm.doc.folder_path || "";
            let rel_dp = current_sharepoint_path;
            let display_path;
            if (rel_dp === '/') {
                display_path = base_dp || '/';
            } else {
                display_path = base_dp.replace(/\/$/, '') + '/' + rel_dp.replace(/^\//, '');
            }
            if (display_path.startsWith('//')) { // Normalize if base_dp was '/'
                display_path = display_path.substring(1);
            }
            if (display_path === "") display_path = "/";


            if (frm.fields_dict['sharepoint_contents_html']) {
                frm.fields_dict['sharepoint_contents_html'].$wrapper.html(`<div class="text-muted"><i class="fa fa-spinner fa-spin"></i> ${__('Loading contents for:')} <strong>${display_path}</strong></div>`);
            }

            frappe.call({
                method: "document_management.document_management.utils.sharepoint_integration.list_sharepoint_folder_contents",
                args: {
                    folder_docname: current_folder_docname,
                    relative_path: current_sharepoint_path // This is relative to the Folder DocType's root
                },
                callback: function(r) {
                    let html = `<p class="text-muted small mb-2">${__('Current Location:')} <strong>${display_path}</strong></p>`;
                    if (r.message && r.message.items) {
                        if (r.message.items.length === 0) {
                            html += '<div class="text-muted">' + __('This folder is empty.') + '</div>';
                        } else {
                            html += '<ul class="list-group">';
                            r.message.items.forEach(item => {
                                let item_style = "text-decoration: none; color: inherit; display: block;";
                                html += `<li class="list-group-item list-group-item-action p-2">
                                            <a href="#" class="sharepoint-item" data-path="${item.path}" data-name="${item.name}" data-is-folder="${item.is_folder}" style="${item_style}">
                                                <i class="fa ${item.is_folder ? 'fa-folder' : 'fa-file-o'} mr-2"></i> ${item.name}
                                            </a>
                                         </li>`;
                            });
                            html += '</ul>';
                        }

                        if (frm.fields_dict['sharepoint_contents_html']) {
                            frm.fields_dict['sharepoint_contents_html'].$wrapper.html(html);
                        }

                        // Handle item clicks
                        frm.get_field('sharepoint_contents_html').$wrapper.find('a.sharepoint-item').on('click', function(e) {
                            e.preventDefault();
                            let clicked_item_path = $(this).data('path');
                            let clicked_item_is_folder = $(this).data('is-folder');
                            let clicked_item_name = $(this).data('name');

                            if (clicked_item_is_folder) {
                                current_sharepoint_path = clicked_item_path;
                                update_sharepoint_contents_display(); // Navigate into folder
                            } else {
                                // File actions
                                // File actions
                                const file_action_dialog = new frappe.ui.Dialog({
                                    title: __("Actions for file: {0}", [clicked_item_name]),
                                    fields: [
                                        {
                                            fieldname: 'info_html',
                                            fieldtype: 'HTML',
                                            options: `<p>${__('Selected file:')} <strong>${clicked_item_name}</strong></p><p class="text-muted small mb-3">${__('Path:')} ${clicked_item_path}</p>`
                                        },
                                        {
                                            fieldname: 'open_link_btn',
                                            fieldtype: 'Button',
                                            label: __('Open File Link'),
                                            btn_class: 'btn-primary',
                                            click: () => {
                                                frappe.call({
                                                    method: 'document_management.document_management.utils.sharepoint_integration.get_sharepoint_item_details',
                                                    args: {
                                                        target_folder_docname: current_folder_docname,
                                                        relative_path_to_item: clicked_item_path
                                                    },
                                                    callback: function(r_item) {
                                                        if (r_item.message && r_item.message.is_file && r_item.message.sharepoint_link) {
                                                            window.open(r_item.message.sharepoint_link, '_blank');
                                                        } else if (r_item.message && r_item.message.error) {
                                                            frappe.msgprint({ title: __('Error'), indicator: 'red', message: r_item.message.error });
                                                        } else if (r_item.exc) {
                                                            frappe.msgprint({ title: __('Server Error'), indicator: 'red', message: __('Error fetching file link. Check server logs.')});
                                                            console.error("Get Item Link Error:", r_item.exc);
                                                        } else {
                                                            frappe.msgprint(__('File link not available or item is not a file.'));
                                                        }
                                                    },
                                                    error: function(err_item) {
                                                        frappe.msgprint({ title: __('Network Error'), indicator: 'red', message: __('Failed to communicate for file link.')});
                                                        console.error("AJAX Error (Get Item Link):", err_item);
                                                    }
                                                });
                                                file_action_dialog.hide();
                                            }
                                        },
                                        {
                                            fieldname: 'download_btn',
                                            fieldtype: 'Button',
                                            label: __('Download File'),
                                            click: () => {
                                                // First, get the item details to fetch the sharepoint_link (teams_link)
                                                frappe.show_alert({ message: __('Fetching file details for download...'), indicator: 'info' });
                                                frappe.call({
                                                    method: 'document_management.document_management.utils.sharepoint_integration.get_sharepoint_item_details',
                                                    args: {
                                                        target_folder_docname: current_folder_docname,
                                                        relative_path_to_item: clicked_item_path
                                                    },
                                                    callback: function(r_item_details_for_download) {
                                                        if (r_item_details_for_download.message && r_item_details_for_download.message.is_file && r_item_details_for_download.message.sharepoint_link) {
                                                            let teams_link_for_download = r_item_details_for_download.message.sharepoint_link;
                                                            
                                                            // Now call the download method
                                                            frappe.show_alert({ message: __('Initiating download...'), indicator: 'info' });
                                                            frappe.call({
                                                                method: 'document_management.document_management.utils.sharepoint_integration.download_items',
                                                                args: {
                                                                    teams_link: teams_link_for_download
                                                                },
                                                                callback: function(r_download) {
                                                                    if (r_download.message && r_download.message.success) {
                                                                        frappe.show_alert({ message: __('Download started. Check your browser downloads.'), indicator: 'green', display_length: 5000 });
                                                                        // The actual download is handled by the browser via content-disposition from server
                                                                    } else if (r_download.message && r_download.message.error) {
                                                                        frappe.msgprint({ title: __('Download Error'), indicator: 'red', message: r_download.message.error });
                                                                    } else if (r_download.exc) {
                                                                        frappe.msgprint({ title: __('Server Error'), indicator: 'red', message: __('An error occurred while trying to download the file. Check server logs.')});
                                                                        console.error("Download File Error:", r_download.exc);
                                                                    } else {
                                                                        // This case might occur if download_items doesn't return a clear success/error message structure
                                                                        // but still initiates a download.
                                                                        frappe.show_alert({ message: __('Download initiated. If it does not start, please check console logs or contact support.'), indicator: 'orange', display_length: 7000 });
                                                                    }
                                                                },
                                                                error: function(err_download) {
                                                                    frappe.msgprint({ title: __('Network Error'), indicator: 'red', message: __('Failed to communicate for file download.')});
                                                                    console.error("AJAX Error (Download File):", err_download);
                                                                }
                                                            });
                                                        } else if (r_item_details_for_download.message && r_item_details_for_download.message.error) {
                                                            frappe.msgprint({ title: __('Error'), indicator: 'red', message: r_item_details_for_download.message.error });
                                                        } else if (r_item_details_for_download.exc) {
                                                            frappe.msgprint({ title: __('Server Error'), indicator: 'red', message: __('Error fetching file details for download. Check server logs.')});
                                                            console.error("Get Item Details for Download Error:", r_item_details_for_download.exc);
                                                        } else {
                                                            frappe.msgprint(__('Could not retrieve file details necessary for download.'));
                                                        }
                                                    },
                                                    error: function(err_item_details) {
                                                        frappe.msgprint({ title: __('Network Error'), indicator: 'red', message: __('Failed to fetch file details for download.')});
                                                        console.error("AJAX Error (Get Item Details for Download):", err_item_details);
                                                    }
                                                });
                                                file_action_dialog.hide(); // Hide dialog after initiating the process
                                            }
                                        },
                                        {
                                            fieldname: 'create_incoming_btn',
                                            fieldtype: 'Button',
                                            label: __('Create Incoming Document'),
                                            click: () => {
                                                let base_nd = frm.doc.folder_path || "";
                                                let rel_nd = clicked_item_path;
                                                let absolute_path_for_new_doc;
                                                if (rel_nd === '/') { absolute_path_for_new_doc = base_nd || '/'; }
                                                else { absolute_path_for_new_doc = (base_nd.replace(/\/$/, '') + '/' + rel_nd.replace(/^\//, '')).replace(/\/\//g, '/'); }
                                                if (absolute_path_for_new_doc === "") absolute_path_for_new_doc = "/";

                                                frappe.call({
                                                    method: 'document_management.document_management.utils.sharepoint_integration.get_sharepoint_item_details',
                                                    args: {
                                                        target_folder_docname: current_folder_docname,
                                                        relative_path_to_item: clicked_item_path // This is relative to the Folder's root
                                                    },
                                                    callback: function(r_item_details) {
                                                        let teams_link = '';
                                                        if (r_item_details.message && r_item_details.message.sharepoint_link) {
                                                            teams_link = r_item_details.message.sharepoint_link;
                                                        } else {
                                                            frappe.throw(__("Could not fetch SharePoint link for the selected file. Document creation aborted. Path: {0}", [clicked_item_path]));
                                                            return; // Stop further execution
                                                        }
                                                        frappe.new_doc('Incoming Document', {
                                                            folder: frm.doc.name,
                                                            path: absolute_path_for_new_doc,
                                                            teams_link: teams_link
                                                        });
                                                        file_action_dialog.hide();
                                                    },
                                                    error: function(err_details) {
                                                        console.error("Error fetching SharePoint link for new Incoming Document: ", err_details);
                                                        frappe.throw(__("Error fetching SharePoint link: {0}. Document creation aborted.", [err_details.message || JSON.stringify(err_details)]));
                                                    }
                                                });
                                            }
                                        },
                                        {
                                            fieldname: 'create_outgoing_btn',
                                            fieldtype: 'Button',
                                            label: __('Create Outgoing Document'),
                                            click: () => {
                                                let base_nd = frm.doc.folder_path || "";
                                                let rel_nd = clicked_item_path;
                                                let absolute_path_for_new_doc;
                                                if (rel_nd === '/') { absolute_path_for_new_doc = base_nd || '/'; }
                                                else { absolute_path_for_new_doc = (base_nd.replace(/\/$/, '') + '/' + rel_nd.replace(/^\//, '')).replace(/\/\//g, '/'); }
                                                if (absolute_path_for_new_doc === "") absolute_path_for_new_doc = "/";

                                                frappe.call({
                                                    method: 'document_management.document_management.utils.sharepoint_integration.get_sharepoint_item_details',
                                                    args: {
                                                        target_folder_docname: current_folder_docname,
                                                        relative_path_to_item: clicked_item_path // Relative to Folder's root
                                                    },
                                                    callback: function(r_item_details) {
                                                        let teams_link = '';
                                                        if (r_item_details.message && r_item_details.message.sharepoint_link) {
                                                            teams_link = r_item_details.message.sharepoint_link;
                                                        } else {
                                                            frappe.throw(__("Could not fetch SharePoint link for the selected file. Document creation aborted. Path: {0}", [clicked_item_path]));
                                                            return; // Stop further execution
                                                        }
                                                        frappe.new_doc('Outgoing Document', {
                                                            folder: frm.doc.name,
                                                            path: absolute_path_for_new_doc,
                                                            teams_link: teams_link
                                                        });
                                                        file_action_dialog.hide();
                                                    },
                                                    error: function(err_details) {
                                                        console.error("Error fetching SharePoint link for new Outgoing Document: ", err_details);
                                                        frappe.throw(__("Error fetching SharePoint link: {0}. Document creation aborted.", [err_details.message || JSON.stringify(err_details)]));
                                                    }
                                                });
                                            }
                                        },
                                        {
                                            fieldname: 'cancel_btn',
                                            fieldtype: 'Button',
                                            label: __('Cancel'),
                                            click: () => {
                                                file_action_dialog.hide();
                                            }
                                        }
                                    ]
                                    // No primary_action or actions array here, buttons are defined as fields.
                                });
                                file_action_dialog.show();
                            }
                        });
                    } else if (r.message && r.message.error) {
                        if (frm.fields_dict['sharepoint_contents_html']) {
                            frm.fields_dict['sharepoint_contents_html'].$wrapper.html(html + `<div class="text-danger mt-2">${__('Error')}: ${r.message.error}</div>`);
                        }
                    } else if (r.exc) {
                        if (frm.fields_dict['sharepoint_contents_html']) {
                            frm.fields_dict['sharepoint_contents_html'].$wrapper.html(html + `<div class="text-danger mt-2">${__('Server error while listing contents.')}</div>`);
                        }
                        console.error("List Contents Error:", r.exc);
                    } else {
                        if (frm.fields_dict['sharepoint_contents_html']) {
                            frm.fields_dict['sharepoint_contents_html'].$wrapper.html(html + '<div class="text-muted mt-2">' + __('Folder is empty or unable to load contents.') + '</div>');
                        }
                    }
                },
                error: function(err) {
                    if (frm.fields_dict['sharepoint_contents_html']) {
                        frm.fields_dict['sharepoint_contents_html'].$wrapper.html(html + `<div class="text-danger mt-2">${__('Network error while listing contents.')}</div>`);
                    }
                    console.error("AJAX Error (List Contents):", err);
                }
            });
        }

        // Trigger update when the form is refreshed
        if (!frm.is_new()) {
            update_sharepoint_contents_display();
        } else {
            // Optionally, clear the HTML field or show a "save first" message if the doc is new
            if (frm.fields_dict['sharepoint_contents_html']) {
                frm.fields_dict['sharepoint_contents_html'].$wrapper.html('<div class="text-muted">' + __('Save the document to see SharePoint contents.') + '</div>');
            }
            if (frm.get_field('sharepoint_contents_tab')) {
                frm.get_field('sharepoint_contents_tab').toggle(false);
            }
        }

        // Optional: Trigger update when the tab is clicked (if needed for performance or initial load)
        // frm.fields_dict['sharepoint_contents_tab'].wrapper.on('click', function() {
        //     update_sharepoint_contents_display();
        // });
    }
});