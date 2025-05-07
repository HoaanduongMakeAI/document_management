# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
import requests

# Import the new utility function (Corrected Path)
from document_management.document_management.utils.sharepoint_integration import get_group_details_and_sharepoint_ids

@frappe.whitelist()
def fetch_sharepoint_ids_from_group():
	"""
	Fetches SharePoint IDs for the singleton DocumentManagementSettings instance
	using the centralized utility function.

	Returns:
		dict: Dictionary containing site_id and drive_id
	"""
	settings = frappe.get_single("Document Management Settings")

	# Basic validation remains here
	if not settings.connected_app:
		frappe.throw("Please select a Connected App first in Document Management Settings")
	if not settings.entra_group_id:
		frappe.throw("Please enter the Microsoft Entra Group ID first in Document Management Settings")

	try:
		# Call the utility function - it now returns group_name, site_id, drive_id
		group_name, site_id, drive_id = get_group_details_and_sharepoint_ids(settings.entra_group_id)

		# Update settings if IDs are successfully retrieved
		if site_id and drive_id:
			settings.sharepoint_site_id = site_id
			settings.sharepoint_drive_id = drive_id
			settings.save() # Save the singleton document
			frappe.msgprint("Successfully fetched and updated SharePoint Site ID and Drive ID.", indicator="green")
			return {
				"sharepoint_site_id": site_id,
				"sharepoint_drive_id": drive_id
			}
		else:
			# This case should ideally be handled by exceptions in the utility function
			frappe.throw("Failed to retrieve SharePoint IDs. The utility function returned empty values.")

	except Exception as e:
		# Error logging and throwing are handled within get_group_details_and_sharepoint_ids
		# We re-throw here to ensure the client-side gets the error message
		frappe.throw(f"Failed to fetch SharePoint IDs: {str(e)}")

class DocumentManagementSettings(Document):
    @frappe.whitelist()
    def fetch_sharepoint_ids_from_group(self):
        """Wrapper method for backward compatibility"""
        return fetch_sharepoint_ids_from_group()