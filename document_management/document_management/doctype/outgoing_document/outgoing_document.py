# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from frappe.utils import get_abbr

# Import the upload function - Adjust path due to directory move
from document_management.document_management.utils.sharepoint_integration import upload_file_to_sharepoint

class OutgoingDocument(Document):
	# autoname is now handled by Naming Series in JSON

	def before_save(self):
		"""
		Before save hook: Trigger SharePoint upload if a file is newly attached
		via the 'document_attachment' field.
		"""
		# Check if the document_attachment field exists and has changed
		if self.has_value_changed("document_attachment") and self.document_attachment:
			frappe.msgprint(f"New attachment '{self.document_attachment}' detected for Outgoing Document '{self.name}'. Attempting SharePoint upload.")

			# Extract the File DocType name from the URL stored in the Attach Image field
			file_doc_name = self.document_attachment.split("/")[-1]

			if not frappe.db.exists("File", file_doc_name):
				frappe.logger.error(f"File DocType '{file_doc_name}' not found for attachment URL '{self.document_attachment}' in Outgoing Document '{self.name}'.")
				return # Exit before attempting upload

			action_details = {
				"action": "Attach/Save",
				"user": frappe.session.user
			}
			try:
				# Call the upload function from the utility module
				upload_result = upload_file_to_sharepoint(self, file_doc_name, action_details)

				if upload_result and upload_result.get("sharepoint_link"):
					# Update the read-only teams_link field with the result
					self.teams_link = upload_result["sharepoint_link"]
					frappe.msgprint(f"File uploaded to SharePoint: {self.teams_link}", indicator="green", alert=True)
				else:
					frappe.logger.warning(f"SharePoint upload for '{file_doc_name}' in Outgoing Document '{self.name}' completed but returned no link or failed silently.", upload_result)

			except Exception as e:
				frappe.logger.error(f"SharePoint upload failed for '{file_doc_name}' in Outgoing Document '{self.name}': {e}")
				frappe.msgprint(f"SharePoint upload failed: {e}", indicator="red", raise_exception=False, alert=True)
				# self.teams_link = None # Optional: Clear link on failure
				# self.document_attachment = None # Optional: Clear attachment on failure

		# Remove old logic
		# if not self.name or "TEMP" in self.name or not self.name.startswith("OUT-"):
		#      self.autoname()
		# if self.base_document:
		#      ...

	# get_latest_attachment is no longer needed

	# TODO: Add workflow state change hooks (on_submit, on_approve, etc.)
	# These hooks could potentially call upload_file_to_sharepoint again if needed
	# for specific workflow actions, passing different action_details.