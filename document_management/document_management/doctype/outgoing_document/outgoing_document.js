frappe.ui.form.on('Outgoing Document', {
    refresh: function(frm) {
        // Add custom button functionality if the document is not new
        if (!frm.is_new()) {
            frm.add_custom_button(__('Attach and Upload to SharePoint'), function() {
                // Use Frappe's file uploader
                frappe.prompt(
                    [
                        {
                            fieldname: 'file_upload',
                            fieldtype: 'Attach',
                            label: __('Upload File'),
                            reqd: 1
                        }
                    ],
                    function(values) {
                        if (values && values.file_upload) {
                            let file_url = values.file_upload;
                            // Extract the file_doc_name from the URL
                            let file_doc_name = file_url.split('/').pop();

                            frappe.show_alert({ message: __('Uploading to SharePoint...'), indicator: 'blue' });

                            // Call the server-side whitelisted function for Outgoing Document
                            frappe.call({
                                method: 'document_management.document_management.doctype.outgoing_document.outgoing_document.upload_outgoing_file_via_modal', // Correct method path
                                args: {
                                    docname: frm.doc.name,
                                    file_doc_name: file_doc_name
                                },
                                callback: function(r) {
                                    if (r.message && r.message.sharepoint_link) {
                                        // Update the teams_link field on the form
                                        frm.set_value('teams_link', r.message.sharepoint_link);
                                        frappe.show_alert({ message: __('File successfully uploaded and linked.'), indicator: 'green' });
                                        frm.save(); // Optionally save the form after successful upload
                                    } else if (r.message && r.message.error) {
                                        frappe.msgprint({
                                            title: __('SharePoint Upload Error'),
                                            indicator: 'red',
                                            message: r.message.error
                                        });
                                    } else if (r.exc) {
                                        // Handle Python exceptions thrown by frappe.throw
                                        frappe.msgprint({
                                            title: __('Server Error'),
                                            indicator: 'red',
                                            message: __('An error occurred during SharePoint upload. Please check server logs.')
                                        });
                                        console.error("SharePoint Upload Error:", r.exc);
                                    } else {
                                         frappe.msgprint({
                                            title: __('Upload Issue'),
                                            indicator: 'orange',
                                            message: __('Upload completed but no link was returned. Please check the document and SharePoint.')
                                        });
                                    }
                                },
                                error: function(r) {
                                     frappe.msgprint({
                                        title: __('Network Error'),
                                        indicator: 'red',
                                        message: __('Failed to communicate with the server for SharePoint upload.')
                                    });
                                    console.error("AJAX Error:", r);
                                }
                            });
                        }
                    },
                    __('Upload File to SharePoint'), // Dialog Title
                    __('Upload') // Primary Button Label
                );
            }, __('Actions')); // Add button to 'Actions' group
        }

        // Make the button field read-only or hide it if needed,
        // as we are using a custom button added via script.
        // frm.set_df_property('upload_to_sharepoint_btn', 'hidden', 1);
    },

    // Example: Clear link if department changes
    // department: function(frm) {
    //     if (frm.doc.teams_link) {
    //         // Logic to potentially clear the link if context changes
    //     }
    // }
});