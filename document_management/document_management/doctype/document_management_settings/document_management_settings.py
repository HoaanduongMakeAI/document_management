# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
import requests

@frappe.whitelist()
def fetch_sharepoint_ids_from_group():
	"""
	Fetches SharePoint IDs for the singleton DocumentManagementSettings instance

	Returns:
		dict: Dictionary containing site_id and drive_id
	"""
	settings = frappe.get_single("Document Management Settings")

	if not settings.connected_app:
		frappe.throw("Please select a Connected App first in Document Management Settings")
	if not settings.entra_group_id:
		frappe.throw("Please enter the Microsoft Entra Group ID first in Document Management Settings")

	try:
		from document_management.document_management.utils.sharepoint_integration import get_headers
		headers = get_headers()

		# Get SharePoint site
		site_url = f"https://graph.microsoft.com/v1.0/groups/{settings.entra_group_id}/sites/root"
		site_response = requests.get(site_url, headers=headers)
		site_response.raise_for_status()
		site_data = site_response.json()
		# frappe.msgprint(f"Site data: {site_data}")
		site_id = site_data.get("id")
		if not site_id:
			frappe.throw("Could not retrieve Site ID from Group information")

		# Get document library
		drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
		drives_response = requests.get(drives_url, headers=headers)
		drives_response.raise_for_status()
		drives_data = drives_response.json()
		drive_id = None
		if drives_data.get("value"):
			for drive in drives_data["value"]:
				if drive.get("name", "").lower() == "documents":
					drive_id = drive.get("id")
					break
			if not drive_id:
				drive_id = drives_data["value"][0].get("id")

		if not drive_id:
			frappe.throw("Could not retrieve Drive ID from Site information")

		# Update settings
		settings.sharepoint_site_id = site_id
		settings.sharepoint_drive_id = drive_id
		settings.save()
		# frappe.db.commit()
		return {
			"sharepoint_site_id": site_id,
			"sharepoint_drive_id": drive_id
		}

	except requests.exceptions.RequestException as e:
		err_msg = e.response.text if e.response else str(e)
		frappe.log_error(f"Graph API Error: {err_msg}")
		frappe.throw(f"Error communicating with Microsoft Graph API: {err_msg}")
	except Exception as e:
		frappe.log_error(f"Error: {str(e)}")
		frappe.throw(f"An unexpected error occurred: {str(e)}")

class DocumentManagementSettings(Document):
    @frappe.whitelist()
    def fetch_sharepoint_ids_from_group(self):
        """Wrapper method for backward compatibility"""
        return fetch_sharepoint_ids_from_group()