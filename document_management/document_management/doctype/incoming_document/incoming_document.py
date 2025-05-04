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

# Whitelisted functions moved to utils/sharepoint_integration.py