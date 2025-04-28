# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from frappe.utils import get_abbr

# Import the upload function - Adjust path due to directory move
from document_management.document_management.utils.sharepoint_integration import upload_file_to_sharepoint

class OutgoingDocument(Document):
    def autoname(self):
        """
        Sets the name using the format: OUT-{abbr_department}-{YYYY}-{#####}
        Requires the base_document to be linked first to get the department.
        """
        dept_abbr = "NA" # Default if department not found
        base_doc = None
        # Ensure base_document is linked before trying to get department
        if self.base_document:
             # Try to fetch base_document if linked but not loaded yet (e.g., during validate)
             try:
                base_doc = frappe.get_doc("Document", self.base_document)
             except Exception:
                 frappe.log_error(f"Could not load Base Document {self.base_document} during autoname for Outgoing Document.")
                 base_doc = None # Ensure base_doc is None if fetch fails

        if not base_doc:
              # Cannot autoname without base document and department yet.
              # Naming will be set later, perhaps in before_insert or before_save
              # For now, let's prevent error but log it.
              frappe.log_warning(f"Cannot generate full name for Outgoing Document yet: Base Document not linked or loaded.")
              # Set a temporary name or rely on manual naming / later hook
              # self.name = frappe.model.naming.make_autoname(f"OUT-TEMP-.YYYY.-#####") # Example temporary
              return # Exit early, name will be set later


        # Proceed if base_doc is available
        if base_doc and base_doc.department:
             dept_info = frappe.get_value("Department", base_doc.department, ["name", "abbr"], as_dict=True)
             if dept_info and dept_info.abbr:
                 dept_abbr = dept_info.abbr
             elif dept_info: # Fallback to department name abbreviation if abbr field is empty
                 dept_abbr = get_abbr(dept_info.name)

        self.abbr_department = dept_abbr # Store abbreviation
        # Only set name if dept_abbr is not NA (meaning department was found)
        if dept_abbr != "NA":
             # Avoid resetting name if it's already set correctly (e.g., during save after validate)
             current_prefix = f"OUT-{dept_abbr}-"
             if not self.name or not self.name.startswith(current_prefix):
                  self.name = frappe.model.naming.make_autoname(f"OUT-{dept_abbr}-.YYYY.-#####")
        # else: name remains unset or temporary


    def validate(self):
        """
        Validate hook: Trigger SharePoint upload if a file is newly attached.
        """
        latest_attachment = self.get_latest_attachment()

        if latest_attachment:
            already_uploaded = False
            if not self.is_new():
                 for version in self.get("versions", []):
                      if version.file_url and version.file_url.endswith(f"/{latest_attachment.name}"):
                           already_uploaded = True
                           frappe.log_info(f"Attachment '{latest_attachment.name}' seems to be already uploaded for Outgoing Document '{self.name}'. Skipping upload.")
                           break

            if not already_uploaded:
                frappe.log_info(f"New attachment '{latest_attachment.name}' detected for Outgoing Document '{self.name}'. Attempting SharePoint upload.")
                action_details = {
                    "action": "Attach/Validate",
                    "user": frappe.session.user
                }
                try:
                    upload_result = upload_file_to_sharepoint(self, latest_attachment.name, action_details)
                    if upload_result:
                        frappe.msgprint(f"File uploaded to SharePoint: {upload_result['sharepoint_link']}", indicator="green", alert=True)
                except Exception as e:
                    frappe.msgprint(f"SharePoint upload failed: {e}", indicator="red", raise_exception=False, alert=True)


    def before_save(self):
        # --- Autonaming Logic (Ensure it runs if name not set) ---
        # If name wasn't set in autoname (e.g., base_doc not ready), try again here
        if not self.name or "TEMP" in self.name or not self.name.startswith("OUT-"):
             self.autoname() # Try setting the name again now that fields might be populated

        # --- Link subject/description ---
        if self.base_document:
             # Load base_doc only if needed
             base_doc_data = {}
             if not self.subject or self.subject == self.name or not self.description:
                  base_doc_data = frappe.db.get_value("Document", self.base_document, ["subject", "description", "department"], as_dict=True)

             if base_doc_data:
                 if not self.subject or self.subject == self.name:
                     if base_doc_data.subject:
                          self.subject = base_doc_data.subject
                 if not self.description:
                     if base_doc_data.description:
                          self.description = base_doc_data.description

                 # Ensure department abbreviation is set for naming/consistency
                 if not self.abbr_department and base_doc_data.department:
                     dept_info = frappe.get_value("Department", base_doc_data.department, ["name", "abbr"], as_dict=True)
                     if dept_info and dept_info.abbr:
                         self.abbr_department = dept_info.abbr
                     elif dept_info:
                         self.abbr_department = get_abbr(dept_info.name)
                     # If name still needs setting after getting abbr
                     if not self.name or "TEMP" in self.name or not self.name.startswith("OUT-"):
                          self.autoname()


    # Remove the placeholder handle_sharepoint_upload method
    # def handle_sharepoint_upload(self):
    #     pass

    def get_latest_attachment(self):
        """Helper to get the most recently attached File DocType."""
        attached_file_name = getattr(self, "_file_name", None)
        if attached_file_name:
             if "/" in attached_file_name:
                  file_doc_name = attached_file_name.split("/")[-1]
                  if frappe.db.exists("File", file_doc_name):
                       return frappe.get_doc("File", file_doc_name)

        attachments = frappe.get_all(
            "File",
            filters={"attached_to_doctype": self.doctype, "attached_to_name": self.name},
            fields=["name", "creation"],
            order_by="creation desc",
            limit=1
        )
        if attachments:
            return frappe.get_doc("File", attachments[0].name)
        return None

    # TODO: Add workflow state change hooks (on_submit, on_approve, etc.)
    # These hooks would call upload_file_to_sharepoint for actions like "Approve", "Issue"
    # passing different action_details and potentially handling the final version.