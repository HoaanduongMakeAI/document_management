frappe.ui.form.on('Incoming Document', {
    custom_handle_attach_and_upload: function(frm) {
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

                    frappe.show_alert({ message: __('Fetching file details...'), indicator: 'blue' });

                    // Step 1: Get the File DocType name (hash) from the file_url
                    frappe.call({
                        method: 'document_management.document_management.utils.sharepoint_integration.get_latest_file_doc_name_by_url',
                        args: {
                            file_url: file_url
                        },
                        callback: function(r) {
                            if (r.message) {
                                let file_doc_name = r.message;

                                if (file_doc_name) {
                                    frappe.show_alert({ message: __('File details fetched. Uploading to SharePoint...'), indicator: 'blue' });

                                    // Step 2: Call the generalized upload function with doctype, docname, and file_doc_name
                                    frappe.call({
                                        method: 'document_management.document_management.utils.sharepoint_integration.upload_file_via_modal',
                                        args: {
                                            doctype: frm.doc.doctype, // Pass the doctype
                                            docname: frm.doc.name,
                                            file_doc_name: file_doc_name // Pass the file hash
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
                                } else {
                                    // Handle case where file_doc_name was not returned
                                    frappe.msgprint({
                                        title: __('File Not Found'),
                                        indicator: 'red',
                                        message: __('Could not find the uploaded file record. Please try uploading again.')
                                    });
                                }
                            } else if (r.exc) {
                                 // Handle Python exceptions from get_latest_file_doc_name_by_url
                                frappe.msgprint({
                                    title: __('Server Error'),
                                    indicator: 'red',
                                    message: __('An error occurred while fetching file details. Please check server logs.')
                                });
                                console.error("Get File Details Error:", r.exc);
                            } else {
                                frappe.msgprint({
                                    title: __('File Details Issue'),
                                    indicator: 'orange',
                                    message: __('Could not retrieve file details. Please try again.')
                                });
                            }
                        },
                        error: function(r) {
                             frappe.msgprint({
                                title: __('Network Error'),
                                indicator: 'red',
                                message: __('Failed to communicate with the server for file details.')
                            });
                            console.error("AJAX Error:", r);
                        }
                    });
                }
            },
            __('Upload File to SharePoint'), // Dialog Title
            __('Upload') // Primary Button Label
        );
    },

    refresh: function(frm) {
        if (!frm.is_new()) {
            // Check if the button field 'upload_to_sharepoint_btn' (defined in JSON) exists
            if (frm.fields_dict['upload_to_sharepoint_btn']) {
                // Set the click handler for the button
                frm.set_df_property('upload_to_sharepoint_btn', 'click', () => {
                    frm.events.custom_handle_attach_and_upload(frm);
                });
                // Ensure the button is visible (it should be by default)
                // frm.set_df_property('upload_to_sharepoint_btn', 'hidden', 0);
            }
        }
    },

    // You might want to clear the teams_link if the document type changes, for example
    // document_type: function(frm) {
    //     if (frm.doc.teams_link) {
    //         // Logic to potentially clear the link if context changes
    //     }
    // }
});