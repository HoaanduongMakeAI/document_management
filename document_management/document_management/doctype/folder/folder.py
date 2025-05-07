# Copyright (c) 2025, makeAI and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
import copy # Import the copy module
# Corrected import path
from document_management.document_management.utils.sharepoint_integration import create_sharepoint_folder_if_not_exists

class Folder(Document):
    def autoname(self):
        """
        Sets the name of the Folder document based on the linked Microsoft Entra Group name and folder path.
        Format: (<Group Name>)<Folder Path>
        """
        # Store a deep copy of the original folder_path
        original_folder_path = copy.deepcopy(self.folder_path)

        if self.microsoft_entra_group and original_folder_path:
            try:
                # Fetch the linked Microsoft Entra Group document
                group_doc = frappe.get_doc("Microsoft Entra Group", self.microsoft_entra_group)
                group_name = group_doc.group_name if group_doc.group_name else "Unnamed Group"
                
                # Ensure folder_path starts with a '/' for consistent formatting
                formatted_folder_path = original_folder_path if original_folder_path.startswith('/') else '/' + original_folder_path

                # Construct the new name
                self.name = f"({group_name}){formatted_folder_path}"
                frappe.msgprint(f"Setting Folder name to: {self.name}", indicator="blue")

            except frappe.DoesNotExistError:
                frappe.log_error(f"Linked Microsoft Entra Group '{self.microsoft_entra_group}' not found for Folder '{original_folder_path}'. Using default naming.", "Folder Naming Error")
                # Fallback to default naming if group doc is not found
                self.name = original_folder_path
            except Exception as e:
                frappe.log_error(f"Error setting Folder name for path '{original_folder_path}': {frappe.get_traceback()}", "Folder Naming Error")
                # Fallback to default naming on other errors
                self.name = original_folder_path
        elif original_folder_path:
             # If no group is linked, use just the folder path as the name
             self.name = original_folder_path
        
        # Ensure folder_path is not overwritten by the naming process
        if original_folder_path and self.folder_path != original_folder_path:
            self.folder_path = original_folder_path


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