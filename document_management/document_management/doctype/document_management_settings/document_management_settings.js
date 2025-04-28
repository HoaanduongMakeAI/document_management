frappe.ui.form.on('Document Management Settings', {
    refresh: function(frm) {
        if (frm.is_new()) return;
        
        frm.add_custom_button(__('Fetch SharePoint IDs'), function() {
            frappe.call({
                method: 'document_management.document_management.doctype.document_management_settings.document_management_settings.fetch_sharepoint_ids_from_group',
                args: {},
                callback: function(r) {
                    if (!r.exc) {
                        frappe.show_alert({
                            message: __('SharePoint IDs fetched successfully'),
                            indicator: 'green'
                        });
                        frm.refresh();
                    }
                },
                freeze: true,
                freeze_message: __('Fetching SharePoint IDs...')
            });
        });
    }
});