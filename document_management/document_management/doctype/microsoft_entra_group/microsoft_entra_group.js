// Copyright (c) 2025,  and contributors
// For license information, please see license.txt

frappe.ui.form.on('Microsoft Entra Group', {
    refresh: function(frm) {
        // The 'fetch_ids' button (defined in JSON) will be handled by the event below
    },

    fetch_ids: function(frm) {
        // This function will be called when the 'Fetch IDs' button (from JSON) is clicked
        if (!frm.doc.microsoft_entra_group_id) {
            frappe.msgprint(__('Please enter the Microsoft Entra (Azure AD) Group ID first.'));
            return;
        }
        frappe.call({
            method: 'document_management.document_management.doctype.microsoft_entra_group.microsoft_entra_group.fetch_group_details',
            args: {
                group_id: frm.doc.microsoft_entra_group_id
            },
            callback: function(r) {
                if (r.message) {
                    frm.set_value('group_name', r.message.group_name);
                    frm.set_value('sharepoint_site_id', r.message.sharepoint_site_id);
                    frm.set_value('sharepoint_drive_id', r.message.sharepoint_drive_id);
                    frappe.msgprint(__('IDs and Group Name fetched successfully.'));
                } else {
                    frappe.msgprint(__('Failed to fetch IDs and Group Name.'));
                }
            },
            freeze: true,
            freeze_message: __('Fetching Group Details...')
        });
    }
});