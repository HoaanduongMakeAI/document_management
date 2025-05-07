frappe.ui.form.on('Document Management Settings', {
    custom_fetch_sharepoint_ids: function(frm) {
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
                    frm.reload_doc();
                }
            },
            freeze: true,
            freeze_message: __('Fetching SharePoint IDs...')
        });
    },

    refresh: function(frm) {
        if (frm.is_new()) return;

        // Add the custom button (dynamic)
        frm.add_custom_button(__('Fetch SharePoint IDs'), function() {
            frm.trigger('custom_fetch_sharepoint_ids');
        });

        // The 'fetch_sharepoint_ids_button' (defined in JSON) will be handled by the event below
    },

    fetch_sharepoint_ids_button: function(frm) {
        // This function will be called when the 'fetch_sharepoint_ids_button' (from JSON) is clicked
        frm.trigger('custom_fetch_sharepoint_ids');
    }
});