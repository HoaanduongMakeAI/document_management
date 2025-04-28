# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from frappe.utils import strip_html, get_abbr

# Import the upload function - Adjust path due to directory move
from document_management.document_management.utils.sharepoint_integration import upload_file_to_sharepoint

class IncomingDocument(Document):
	# autoname is now handled by Naming Series in JSON

	def before_save(self):
		"""
		Before save hook: Trigger SharePoint upload if a file is newly attached
		via the 'document_attachment' field.
		"""
		# Check if the document_attachment field exists and has changed
		if self.has_value_changed("document_attachment") and self.document_attachment:
			frappe.log_info(f"New attachment '{self.document_attachment}' detected for Incoming Document '{self.name}'. Attempting SharePoint upload.")

			# Extract the File DocType name from the URL stored in the Attach Image field
			file_doc_name = self.document_attachment.split("/")[-1]

			if not frappe.db.exists("File", file_doc_name):
				frappe.log_error(f"File DocType '{file_doc_name}' not found for attachment URL '{self.document_attachment}' in Incoming Document '{self.name}'.")
				# Optionally clear the attachment field or throw an error
				# self.document_attachment = None
				# frappe.throw(f"Attached file record not found: {file_doc_name}")
				return # Exit before attempting upload

			action_details = {
				"action": "Attach/Save",
				"user": frappe.session.user
			}
			try:
				# Call the upload function from the utility module
				# Pass the File DocType name
				upload_result = upload_file_to_sharepoint(self, file_doc_name, action_details)

				if upload_result and upload_result.get("sharepoint_link"):
					# Update the read-only teams_link field with the result
					self.teams_link = upload_result["sharepoint_link"]
					frappe.msgprint(f"File uploaded to SharePoint: {self.teams_link}", indicator="green", alert=True)
					# The upload function might handle appending to the child table (versions) if needed
				else:
					# Handle cases where upload might succeed but doesn't return a link
					frappe.log_warning(f"SharePoint upload for '{file_doc_name}' in Incoming Document '{self.name}' completed but returned no link or failed silently.", upload_result)
					# Decide if self.teams_link should be cleared or kept as is
					# self.teams_link = None

			except Exception as e:
				# Error is logged and potentially thrown within upload_file_to_sharepoint
				# Log it here as well for context
				frappe.log_error(f"SharePoint upload failed for '{file_doc_name}' in Incoming Document '{self.name}': {e}")
				frappe.msgprint(f"SharePoint upload failed: {e}", indicator="red", raise_exception=False, alert=True)
				# Consider clearing the link if upload fails definitively
				# self.teams_link = None
				# Optionally, clear the attachment field so user has to re-attach
				# self.document_attachment = None

		# Remove old logic related to base_document
		# if self.base_document:
		#      ...

	# get_latest_attachment is no longer needed as we use the specific field

	# TODO: Add workflow state change hooks (e.g., on_submit, on_update_after_submit)
	# These hooks could potentially call upload_file_to_sharepoint again if needed
	# for specific workflow actions, passing different action_details.