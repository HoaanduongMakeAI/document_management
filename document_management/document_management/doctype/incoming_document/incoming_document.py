# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from frappe.utils import strip_html, get_abbr

# Import the upload function - Adjust path due to directory move
from document_management.document_management.utils.sharepoint_integration import upload_file_to_sharepoint

class IncomingDocument(Document):
	# autoname is now handled by Naming Series in JSON

	# before_save hook is removed. Upload logic is now triggered manually via upload_file_via_modal.
	pass

# Whitelisted function to be called from the client-side script
@frappe.whitelist()
def upload_file_via_modal(docname, file_doc_name):
	"""
	Uploads a file (already uploaded to Frappe's File doctype) to SharePoint
	and updates the Incoming Document's teams_link field.

	:param docname: Name of the Incoming Document record.
	:param file_doc_name: Name of the File record (already created by Frappe's uploader).
	"""
	try:
		# Get the Incoming Document
		doc = frappe.get_doc("Incoming Document", docname)

		# Verify the File doctype exists
		if not frappe.db.exists("File", file_doc_name):
			frappe.throw(f"File record '{file_doc_name}' not found. Upload aborted.")
			return {"error": f"File record '{file_doc_name}' not found."}

		action_details = {
			"action": "Upload via Modal",
			"user": frappe.session.user
		}

		# Call the existing SharePoint upload utility function
		upload_result = upload_file_to_sharepoint(doc, file_doc_name, action_details)

		if upload_result and upload_result.get("sharepoint_link"):
			sharepoint_link = upload_result["sharepoint_link"]
			# Update the teams_link field on the document
			doc.db_set("teams_link", sharepoint_link, update_modified=False) # Use db_set to avoid triggering save hooks again
			frappe.msgprint(f"File uploaded to SharePoint: {sharepoint_link}", indicator="green", alert=True)
			return {"sharepoint_link": sharepoint_link}
		else:
			error_msg = f"SharePoint upload for '{file_doc_name}' completed but returned no link or failed. Result: {upload_result}"
			frappe.msgprint(error_msg, indicator="orange", alert=True)
			return {"error": error_msg}

	except Exception as e:
		error_msg = f"SharePoint upload failed for file '{file_doc_name}' on document '{docname}': {e}"
		frappe.log_error(frappe.get_traceback(), f"SharePoint Upload Error (Incoming Document: {docname})")
		frappe.throw(error_msg) # Throw to notify client-side of failure
		# The return below might not be reached due to throw, but included for completeness
		return {"error": str(e)}

	# TODO: Add workflow state change hooks if needed for other actions