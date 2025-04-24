# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from frappe.utils import strip_html, get_abbr

# Import the upload function - Adjust path due to directory move
from ..utils.sharepoint_integration import upload_file_to_sharepoint

class IncomingDocument(Document):
    def autoname(self):
        """
        Sets the name using the format: IN-{sender_abbr}-{YYYY}-{#####}
        """
        # Ensure sender is present for autonaming
        if not self.sender:
             frappe.throw("Sender is required to name the document.")
        sender_abbr = get_abbr(self.sender)
        self.sender_abbr = sender_abbr # Store abbreviation for potential use
        # Use standard naming series generation
        self.name = frappe.model.naming.make_autoname(f"IN-{sender_abbr}-.YYYY.-#####")

    def validate(self):
        """
        Validate hook: Trigger SharePoint upload if a file is newly attached.
        """
        # Check if a file was just attached via the UI (_file_name is a temporary field)
        # or if it's a new document with an attachment field populated (less common without UI)
        # Use get_latest_attachment() helper for robustness
        latest_attachment = self.get_latest_attachment()

        if latest_attachment:
            # Basic check to avoid re-uploading the *exact same file attachment reference*
            # on every save. This doesn't prevent uploading a new file with the same name.
            already_uploaded = False
            if not self.is_new():
                 for version in self.get("versions", []):
                      # Check if a version exists linked to this specific File DocType name
                      if version.file_url and version.file_url.endswith(f"/{latest_attachment.name}"):
                           already_uploaded = True
                           frappe.log_info(f"Attachment '{latest_attachment.name}' seems to be already uploaded for Incoming Document '{self.name}'. Skipping upload.")
                           break

            if not already_uploaded:
                frappe.log_info(f"New attachment '{latest_attachment.name}' detected for Incoming Document '{self.name}'. Attempting SharePoint upload.")
                action_details = {
                    "action": "Attach/Validate",
                    "user": frappe.session.user
                }
                try:
                    # Call the upload function from the utility module
                    # Pass the File DocType name
                    upload_result = upload_file_to_sharepoint(self, latest_attachment.name, action_details)
                    if upload_result:
                        frappe.msgprint(f"File uploaded to SharePoint: {upload_result['sharepoint_link']}", indicator="green", alert=True)
                    # The upload function now handles appending to the child table and saving the doc
                except Exception as e:
                    # Error is logged and thrown within upload_file_to_sharepoint
                    frappe.msgprint(f"SharePoint upload failed: {e}", indicator="red", raise_exception=False, alert=True) # Show error but allow saving for now
            # else: # Logging moved inside the loop for clarity
            #      frappe.log_info(f"Attachment '{latest_attachment.name}' seems to be already uploaded for Incoming Document '{self.name}'. Skipping upload.")


    def before_save(self):
        # Link subject and description from Base Document if not set directly
        # This should run after potential base_document linking logic
        if self.base_document:
             # Check if subject/description are empty or placeholders before overwriting
             if not self.subject or self.subject == self.name:
                 base_subject = frappe.db.get_value("Document", self.base_document, "subject")
                 if base_subject:
                      self.subject = base_subject
             if not self.description:
                 base_description = frappe.db.get_value("Document", self.base_document, "description")
                 if base_description:
                      self.description = base_description


    # Remove the placeholder handle_sharepoint_upload method
    # def handle_sharepoint_upload(self):
    #     pass

    def get_latest_attachment(self):
        """Helper to get the most recently attached File DocType."""
        # Check the temporary field first (set by UI on attach)
        attached_file_name = getattr(self, "_file_name", None)
        if attached_file_name:
             # Frappe often stores the full URL here, extract the File name
             if "/" in attached_file_name:
                  file_doc_name = attached_file_name.split("/")[-1]
                  if frappe.db.exists("File", file_doc_name):
                       return frappe.get_doc("File", file_doc_name)

        # Fallback: Check attachments linked to this document
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


    # TODO: Add workflow state change hooks (e.g., on_submit, on_update_after_submit)
    # These hooks would call upload_file_to_sharepoint for actions like "Approve", "Process"
    # passing different action_details.