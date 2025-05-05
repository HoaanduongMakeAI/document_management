# Copyright (c) 2025, makeAI and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
# Corrected import path
from document_management.document_management.utils.sharepoint_integration import create_sharepoint_folder_if_not_exists

class Folder(Document):
    # This method is called after the document is saved (created or updated)
    def on_update(self):
        self.ensure_sharepoint_folder_exists()

    # on_submit could also be used if folders should only be created upon submission
    # def on_submit(self):
    #    self.ensure_sharepoint_folder_exists()

    def ensure_sharepoint_folder_exists(self):
        if not self.folder_path or not self.microsoft_entra_group:
            frappe.log_error("Folder path or Microsoft Entra Group missing, cannot ensure SharePoint folder.", "Folder SharePoint Sync")
            return
        
        # Check if folder path is just "/" or empty after potential whitespace stripping
        if not self.folder_path.strip() or self.folder_path.strip() == "/":
             frappe.msgprint("Folder path is invalid or points to root, skipping SharePoint folder creation.", indicator="orange")
             return

        try:
            # Fetch the linked Microsoft Entra Group document to get the Drive ID
            group_doc = frappe.get_doc("Microsoft Entra Group", self.microsoft_entra_group)
            sharepoint_drive_id = group_doc.sharepoint_drive_id

            if not sharepoint_drive_id:
                frappe.throw(f"SharePoint Drive ID not found in the linked Microsoft Entra Group '{self.microsoft_entra_group}'. Cannot create folder.")

            # Use the folder_path directly, the utility function handles splitting by '/'
            folder_path_to_create = self.folder_path 

            frappe.msgprint(f"Ensuring SharePoint folder exists for path: '{folder_path_to_create}' in Drive ID: {sharepoint_drive_id}")

            # Call the utility function to create the folder if it doesn't exist
            # This function handles nested creation and checks existence
            folder_id = create_sharepoint_folder_if_not_exists(sharepoint_drive_id, folder_path_to_create)

            if folder_id:
                frappe.msgprint(f"Successfully ensured SharePoint folder exists for '{self.name}'. Folder ID: {folder_id}", indicator="green")
            else:
                # The utility function should throw an error, but add a fallback message
                frappe.msgprint(f"Failed to ensure SharePoint folder exists for '{self.name}'. Check logs for details.", indicator="red")

        except Exception as e:
            frappe.log_error(
                message=f"Error ensuring SharePoint folder for Folder '{self.name}' (Path: {self.folder_path}): {frappe.get_traceback()}",
                title="Folder SharePoint Sync Error"
            )
            # Don't throw here to prevent save/update failure, just log and notify
            frappe.msgprint(f"Error creating/checking SharePoint folder for '{self.name}': {str(e)}", indicator="red", alert=True)