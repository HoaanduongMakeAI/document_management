# Copyright (c) 2025, makeAI and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
# Import the centralized utility function
from document_management.document_management.document_management.utils.sharepoint_integration import get_sharepoint_site_and_drive_ids_for_group

class MicrosoftEntraGroup(Document):
    # This method is called before the document is saved
    def validate(self):
        # Fetch only if group ID is provided and IDs are missing or explicitly requested to refresh
        # Added a check to see if it's a new document or if IDs are missing
        if self.microsoft_entra_group_id and (self.is_new() or not self.sharepoint_site_id or not self.sharepoint_drive_id):
             # Check if connected app is set in settings before proceeding
            settings = frappe.get_cached_doc("Document Management Settings")
            if not settings or not settings.connected_app:
                 frappe.throw("Microsoft Graph Connected App not set in Document Management Settings. Cannot fetch SharePoint IDs.")
            
            self.fetch_and_set_sharepoint_ids()

    # Separated logic for clarity
    def fetch_and_set_sharepoint_ids(self):
        frappe.msgprint(f"Attempting to fetch SharePoint IDs for Group: {self.microsoft_entra_group_id}...")
        try:
            # Call the utility function to get IDs
            site_id, drive_id = get_sharepoint_site_and_drive_ids_for_group(self.microsoft_entra_group_id)

            if site_id and drive_id:
                # Update the document fields directly.
                # These changes will be part of the save transaction.
                self.sharepoint_site_id = site_id
                self.sharepoint_drive_id = drive_id
                frappe.msgprint(f"Successfully fetched and updated SharePoint IDs for Group {self.microsoft_entra_group_id}.", indicator="green")
                # No need to call self.save() here, as validate is part of the save cycle
            else:
                # This case should ideally be handled by exceptions in the utility function
                frappe.throw(f"Could not retrieve valid SharePoint Site ID or Drive ID for Group {self.microsoft_entra_group_id}. Utility function returned empty values.")

        except Exception as e:
            # Error logging and throwing are handled within the utility function
            # Re-throw the error message provided by the utility function
            frappe.throw(f"{str(e)}")

# Note: No need for on_update if logic is in validate, as validate runs before save.