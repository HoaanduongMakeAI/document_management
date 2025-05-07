# Copyright (c) 2025, makeAI and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
import re # For parsing folder_path
# Corrected import path
from document_management.document_management.utils.sharepoint_integration import create_sharepoint_folder_if_not_exists

class Folder(Document):
    def autoname(self):
        """
        Sets the name of the Folder document based on the linked Microsoft Entra Group name and folder path.
        Ensures self.folder_path stores only the path segments (e.g., "General/Công văn").
        Format for self.name: (<Group Name>)/<Folder Path Segments>
        """
        input_folder_path_value = self.folder_path # Value from the form or previous state

        # 1. Determine the actual path component, stripping any potential group prefix from input
        actual_path_segments = input_folder_path_value
        if input_folder_path_value:
            # Attempt to strip a "(group)"-like prefix
            # Check if input starts with '(' and find the corresponding ')'
            idx_closing_paren = -1
            if input_folder_path_value.strip().startswith("("):
                idx_closing_paren = input_folder_path_value.find(")")
            
            if idx_closing_paren != -1:
                # Found a potential group prefix, take the part after ')'
                path_candidate = input_folder_path_value[idx_closing_paren+1:].strip()
                # Remove leading slash if present from this candidate
                actual_path_segments = path_candidate.lstrip('/').strip()
            else:
                # No group-like prefix found, treat the whole thing as path segments
                actual_path_segments = input_folder_path_value.strip()
        
        # Clean the actual_path_segments: remove leading/trailing slashes for consistency
        if actual_path_segments: # Check if not empty string
            actual_path_segments = actual_path_segments.strip('/')
        else: # If actual_path_segments became empty (e.g. input was "(Group)/" or just "/")
            actual_path_segments = "" # Ensure it's an empty string, not None

        # Update self.folder_path to be the clean path segments.
        # This is the critical fix for the folder_path field itself.
        self.folder_path = actual_path_segments

        # 2. Determine the group name for naming
        group_name_for_naming = "Unnamed Group" # Default
        if self.microsoft_entra_group:
            frappe.msgprint(f"DEBUG: self.microsoft_entra_group = {self.microsoft_entra_group}", title="Folder Autoname Debug")
            try:
                group_doc = frappe.get_doc("Microsoft Entra Group", self.microsoft_entra_group)
                frappe.msgprint(f"DEBUG: Fetched group_doc object. Name: {group_doc.name}", title="Folder Autoname Debug")
                frappe.msgprint(f"DEBUG: group_doc.__dict__ = {group_doc.__dict__}", title="Folder Autoname Debug")
                
                retrieved_group_name = group_doc.group_name # Access as attribute
                frappe.msgprint(f"DEBUG: group_doc.group_name (attribute) = {retrieved_group_name}", title="Folder Autoname Debug")

                # Ensure group_name exists and is not just whitespace
                if retrieved_group_name and isinstance(retrieved_group_name, str) and retrieved_group_name.strip():
                    group_name_for_naming = retrieved_group_name.strip()
                    frappe.msgprint(f"DEBUG: Using group_name_for_naming = {group_name_for_naming}", title="Folder Autoname Debug")
                else:
                    frappe.msgprint(f"DEBUG: group_doc.group_name is None, not a string, or empty/whitespace. Using default: '{group_name_for_naming}'", title="Folder Autoname Debug")
            except frappe.DoesNotExistError:
                frappe.msgprint(f"DEBUG: Microsoft Entra Group '{self.microsoft_entra_group}' not found. Using default group name '{group_name_for_naming}'.", title="Folder Autoname Debug", indicator="orange")
                frappe.log_error(f"Linked Microsoft Entra Group '{self.microsoft_entra_group}' not found. Using default group name '{group_name_for_naming}'.", "Folder Naming Error")
            except Exception as e: # Catch any other error during group_doc fetching
                frappe.msgprint(f"DEBUG: Error fetching group name for '{self.microsoft_entra_group}'. Error: {str(e)}. Using default: '{group_name_for_naming}'.", title="Folder Autoname Debug", indicator="red")
                frappe.log_error(f"Error fetching group name for '{self.microsoft_entra_group}'. Using default group name '{group_name_for_naming}'. Details: {frappe.get_traceback()}", "Folder Naming Error")
        else:
            frappe.msgprint("DEBUG: self.microsoft_entra_group is not set. Using default group name.", title="Folder Autoname Debug", indicator="orange")

        # 3. Construct self.name
        if not self.folder_path: # After cleaning, if folder_path is empty
            # Name will be just the group identifier, e.g., "(RealGroup)" or "(Unnamed Group)"
            self.name = f"({group_name_for_naming})"
            # Optional: Log or msgprint this specific case for clarity during operation
            # frappe.msgprint(f"Folder path component is empty. Setting name to group: '{self.name}'. Original input: '{input_folder_path_value}'", indicator="orange", alert=True)
        else:
            # Name will be "(GroupName)/PathSegments"
            self.name = f"({group_name_for_naming})/{self.folder_path}"

        # The user-provided log "Setting Folder name to: {self.name}" can be re-enabled for debugging if needed
        # frappe.msgprint(f"Processed autoname: self.name = '{self.name}', self.folder_path = '{self.folder_path}'", indicator="blue")

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