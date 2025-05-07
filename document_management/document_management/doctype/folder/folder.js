frappe.ui.form.on('Folder', {
    refresh: function(frm) {
        // Initialize current_sharepoint_path relative to the Folder doctype's path
        let current_sharepoint_path = frm.doc.folder_path || '/';
        let current_folder_docname = frm.doc.name; // The Folder doctype's name is the base

        // Add a "Back" button above the HTML field
        if (!frm.custom_back_button_added) {
            frm.add_custom_button(__('Back to Parent Folder'), function() {
                if (current_sharepoint_path && current_sharepoint_path !== frm.doc.folder_path) {
                    let parts = current_sharepoint_path.split('/').filter(p => p.trim() !== '');
                    parts.pop(); // Go up one level
                    current_sharepoint_path = '/' + parts.join('/');
                    if (current_sharepoint_path === '//') current_sharepoint_path = '/';
                    if (current_sharepoint_path === '/') current_sharepoint_path = frm.doc.folder_path; // Go back to the Folder doctype's root path
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

            let display_path = current_sharepoint_path === frm.doc.folder_path ? frm.doc.folder_path : `${frm.doc.folder_path}${current_sharepoint_path.substring(frm.doc.folder_path.length)}`;

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
                                frappe.confirm(
                                    __("Choose an action for file '{0}':", [clicked_item_name]),
                                    [
                                        {
                                            label: __('Open File Link'),
                                            handler: function() {
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
                                            }
                                        },
                                        {
                                            label: __('Create Incoming Document'),
                                            handler: function() {
                                                frappe.new_doc('Incoming Document', {
                                                    folder: frm.doc.name, // Link to this Folder doctype
                                                    path: clicked_item_path, // Set the SharePoint path
                                                    // You might want to pre-fill other fields if possible, e.g., subject from item_name
                                                });
                                            }
                                        },
                                        {
                                            label: __('Create Outgoing Document'),
                                            handler: function() {
                                                 frappe.new_doc('Outgoing Document', {
                                                    folder: frm.doc.name, // Link to this Folder doctype
                                                    path: clicked_item_path, // Set the SharePoint path
                                                    // You might want to pre-fill other fields if possible, e.g., subject from item_name
                                                });
                                            }
                                        }
                                    ],
                                    __('File Actions')
                                );
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
        update_sharepoint_contents_display();

        // Optional: Trigger update when the tab is clicked (if needed for performance or initial load)
        // frm.fields_dict['sharepoint_contents_tab'].wrapper.on('click', function() {
        //     update_sharepoint_contents_display();
        // });
    }
});