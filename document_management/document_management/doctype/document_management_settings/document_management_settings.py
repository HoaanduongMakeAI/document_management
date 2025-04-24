# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
import requests # Needed for Graph API calls

# Import helper from utils module to get authenticated headers
# Use absolute import path
from document_management.document_management.utils.sharepoint_integration import get_headers as get_graph_api_headers

class DocumentManagementSettings(Document):
	# This is a Singleton DocType. Add validation or hooks if needed.
	def validate(self):
		# Example validation: Check if the selected Connected App is valid
		if self.connected_app:
			if not frappe.db.exists("Connected App", self.connected_app):
				frappe.throw(f"Selected Connected App '{self.connected_app}' does not exist.")
			# Could add further checks, e.g., ensure it's a Microsoft Graph app type if possible

	@frappe.whitelist()
	def fetch_sharepoint_ids_from_group(self):
		"""
		Fetches SharePoint Site ID and default Drive ID using the Microsoft Graph API
		based on the provided Entra Group ID and updates the settings document.
		"""
		if not self.connected_app:
			frappe.msgprint("Please select a Connected App first.", indicator="orange", alert=True)
			return

		if not self.entra_group_id:
			frappe.msgprint("Please enter the Microsoft Entra Group ID first.", indicator="orange", alert=True)
			return

		frappe.msgprint("Fetching SharePoint IDs...", indicator="blue")

		try:
			headers = get_graph_api_headers() # Get Bearer token via Connected App

			# 1. Get the default SharePoint site for the group
			site_url = f"https://graph.microsoft.com/v1.0/groups/{self.entra_group_id}/sites/root"
			site_response = requests.get(site_url, headers=headers)
			site_response.raise_for_status() # Raise exception for HTTP errors
			site_data = site_response.json()

			site_id = site_data.get("id")
			if not site_id:
				frappe.throw("Could not retrieve Site ID from Group information.")

			# 2. Get the drives (document libraries) within that site
			drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
			drives_response = requests.get(drives_url, headers=headers)
			drives_response.raise_for_status()
			drives_data = drives_response.json()

			# Assume the first drive is the default document library.
			# A more robust approach might filter by name ('Documents') or webUrl.
			drive_id = None
			if drives_data.get("value"):
				# Prioritize drive named 'Documents' if found
				for drive in drives_data["value"]:
					if drive.get("name", "").lower() == "documents":
						drive_id = drive.get("id")
						break
				# Fallback to the first drive if 'Documents' not found
				if not drive_id:
					drive_id = drives_data["value"][0].get("id")

			if not drive_id:
				frappe.throw("Could not retrieve Drive ID from Site information.")

			# 3. Update the settings document directly in the database
			frappe.db.set_value("Document Management Settings", self.name, {
				"sharepoint_site_id": site_id,
				"sharepoint_drive_id": drive_id
			})
			frappe.db.commit() # Ensure changes are saved

			# Update the current instance's values as well for UI consistency if needed immediately
			self.sharepoint_site_id = site_id
			self.sharepoint_drive_id = drive_id

			frappe.msgprint(f"Successfully fetched and updated IDs:<br>Site ID: {site_id}<br>Drive ID: {drive_id}", indicator="green", alert=True)
			# Optional: Trigger a client-side refresh if needed
			# frappe.local.response['refresh'] = True

		except requests.exceptions.RequestException as e:
			err_msg = e.response.text if e.response else str(e)
			frappe.log_error(f"Graph API Error fetching SharePoint IDs: {err_msg}")
			frappe.msgprint(f"Error communicating with Microsoft Graph API: {err_msg}", indicator="red", alert=True)
		except Exception as e:
			frappe.log_error(f"Error fetching SharePoint IDs: {e}", traceback=True)
			frappe.msgprint(f"An unexpected error occurred: {e}", indicator="red", alert=True)


# The get_settings function has been moved to document_management.utils