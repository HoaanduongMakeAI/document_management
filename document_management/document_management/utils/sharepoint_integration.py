# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
import requests
import os
from frappe.utils import get_site_path, get_files_path, encode
# Import the helper function to get settings
from document_management.document_management.utils import get_settings # Updated import path

@frappe.whitelist()
def get_latest_file_doc_name_by_url(file_url):
	"""
	Finds the most recently created File document matching the given file_url.

	:param file_url: The file_url from the Attach field.
	:return: The name (hash) of the File document or None.
	"""
	if not file_url:
		frappe.log_error("get_latest_file_doc_name_by_url called with empty file_url", "SharePoint Integration")
		return None

	try:
		file_doc_name = frappe.db.get_value(
			"File",
			filters={"file_url": file_url},
			fieldname="name",
			order_by="creation desc",
		)
		if not file_doc_name:
			frappe.log_error(f"No File found for file_url: {file_url}", "SharePoint Integration")
		return file_doc_name
	except Exception as e:
		frappe.log_error(f"Error fetching File for url {file_url}: {e}", "SharePoint Integration")
		return None
@frappe.whitelist()
def upload_file_via_modal(doctype, docname, file_doc_name):
	"""
	Uploads a file (already uploaded to Frappe's File doctype) to SharePoint
	and updates the specified document's teams_link field.

	:param doctype: The DocType of the document to update (e.g., "Incoming Document").
	:param docname: Name of the document record.
	:param file_doc_name: Name of the File record (hash).
	"""
	if not doctype or not docname or not file_doc_name:
		frappe.throw("Missing required arguments: doctype, docname, or file_doc_name.")
		return {"error": "Missing required arguments."}

	try:
		# Get the document dynamically
		doc = frappe.get_doc(doctype, docname)

		# Verify the File doctype exists
		if not frappe.db.exists("File", file_doc_name):
			frappe.throw(f"File record '{file_doc_name}' not found. Upload aborted.")
			# No need for return here as throw stops execution

		action_details = {
			"action": "Upload via Modal",
			"user": frappe.session.user
		}

		# Call the existing SharePoint upload utility function
		# This function already handles folder creation and versioning
		upload_result = upload_file_to_sharepoint(doc, file_doc_name, action_details)

		if upload_result and upload_result.get("sharepoint_link"):
			sharepoint_link = upload_result["sharepoint_link"]
			# Update the teams_link field on the document
			# Use db_set to avoid triggering save hooks again and ensure update
			frappe.db.set_value(doctype, docname, "teams_link", sharepoint_link, update_modified=False)
			frappe.msgprint(f"File uploaded to SharePoint: {sharepoint_link}", indicator="green", alert=True)
			return {"sharepoint_link": sharepoint_link}
		else:
			error_msg = f"SharePoint upload for '{file_doc_name}' completed but returned no link or failed. Result: {upload_result}"
			frappe.msgprint(error_msg, indicator="orange", alert=True)
			# Log the error as well for backend visibility
			frappe.log_error(error_msg, f"SharePoint Upload Issue ({doctype}: {docname})")
			return {"error": error_msg}

	except Exception as e:
		error_msg = f"SharePoint upload failed for file '{file_doc_name}' on document '{doctype} {docname}': {e}"
		frappe.log_error(frappe.get_traceback(), f"SharePoint Upload Error ({doctype}: {docname})")
		frappe.throw(error_msg) # Throw to notify client-side of failure
		# The return below might not be reached due to throw
		# return {"error": str(e)}
def get_sharepoint_settings():
    """Wrapper to get and validate SharePoint settings."""
    # settings = get_settings() # Fetches the singleton DocType instance
    settings = frappe.get_cached_doc("Document Management Settings")
    if not settings:
         frappe.throw("Document Management Settings not found. Please configure them first.")
    if not settings.connected_app:
        frappe.throw("Microsoft Graph Connected App not set in Document Management Settings.")
    if not settings.entra_group_id:
        frappe.throw("Microsoft Entra Group ID not set in Document Management Settings.")
    # if not settings.sharepoint_drive_id:
    #     frappe.throw("SharePoint Drive ID not set in Document Management Settings.")
    return settings

def get_access_token():
    """
    Retrieves the access token using the Connected App specified in settings.
    """
    settings = get_sharepoint_settings()
    try:
    	# Get the Connected App document
    	connected_app = frappe.get_doc("Connected App", settings.connected_app)
    	# Get the active token cache (handles refresh)
    	# Assuming the token is needed for the current user initiating the action
    	token_cache = connected_app.get_active_token(user=frappe.session.user)
   
    	if not token_cache:
    		frappe.throw(f"Could not retrieve token cache for Connected App '{settings.connected_app}' and user '{frappe.session.user}'. Please check configuration and authorization.")
   
    	access_token = token_cache.get_password("access_token")
   
    	if not access_token:
    		frappe.throw(f"Could not retrieve access token from cache for Connected App '{settings.connected_app}'.")
   
    	return access_token
    except Exception as e:
    	# Removed redundant log_error before throw
    	frappe.throw(f"Failed to get access token from Connected App '{settings.connected_app}': {e}")

def get_headers():
    """Returns the authorization headers for Graph API calls."""
    token = get_access_token()
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

def create_sharepoint_folder_if_not_exists(drive_id, folder_path):
    """
    Checks if a folder exists at the specified path within a drive, creates it if not.
    Handles nested folder creation.

    Args:
        drive_id (str): The ID of the SharePoint Drive (Document Library).
        folder_path (str): The relative path of the folder from the drive root (e.g., "FolderA/SubFolderB").

    Returns:
        str: The ID of the folder (existing or newly created).
        None: If an error occurs.
    """
    headers = get_headers()
    segments = folder_path.strip('/').split('/')
    current_path = ""
    parent_item_id = "root" # Start from the drive's root

    for segment in segments:
        if not segment: continue
        current_path = f"{current_path}/{segment}" if current_path else segment
        encoded_segment = encode(segment)
        # Graph API URL to get/create a folder within another item (parent)
        # Using item-id based addressing is more robust than path-based for creation
        folder_check_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent_item_id}/children?$filter=name eq '{encoded_segment}'"

        try:
            frappe.msgprint(f"Checking for SharePoint folder segment: '{segment}'...")
            response = requests.get(folder_check_url, headers=headers)
            response.raise_for_status()
            data = response.json()

            if data.get("value"): # Folder exists
                parent_item_id = data["value"][0]["id"]
                frappe.msgprint(f"SharePoint folder segment '{segment}' exists.")
                # frappe.log_info(f"SharePoint folder segment '{segment}' exists with ID: {parent_item_id}") # Keep log for detailed ID if needed
            else: # Folder does not exist, create it
                frappe.msgprint(f"SharePoint folder segment '{segment}' not found. Creating...")
                create_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent_item_id}/children"
                folder_data = {
                    "name": segment,
                    "folder": {},
                    "@microsoft.graph.conflictBehavior": "fail" # Or 'rename' or 'replace'
                }
                create_response = requests.post(create_url, headers=headers, json=folder_data)
                create_response.raise_for_status()
                new_folder_data = create_response.json()
                parent_item_id = new_folder_data["id"]
                frappe.msgprint(f"Created SharePoint folder segment '{segment}'.")
                # frappe.log_info(f"Created SharePoint folder segment '{segment}' with ID: {parent_item_id}") # Keep log for detailed ID if needed

        except requests.exceptions.RequestException as e:
            err_msg = e.response.text if e.response else str(e)
            # Removed redundant log_error before throw
            frappe.throw(f"Failed to ensure SharePoint folder structure exists: {err_msg}")
            return None
        except Exception as e:
             # Removed redundant log_error before throw
             frappe.throw(f"Unexpected error ensuring SharePoint folder structure: {e}")
             return None

    return parent_item_id # Return the ID of the final folder in the path


def upload_file_to_sharepoint(doc, file_doc_name, action_details):
    """
    Uploads a file attached to a Frappe document to a structured SharePoint folder using settings DocType.

    Args:
        doc (Document): The Frappe document (e.g., IncomingDocument, OutgoingDocument).
        file_doc_name (str): The name of the File DocType record associated with the upload.
        action_details (dict): Details about the action triggering the version (e.g., {'action': 'Submit', 'user': 'user@example.com'}).

    Returns:
        dict: {'sharepoint_link': '...', 'version_id': '...'} or None if upload fails.
    """
    settings = get_sharepoint_settings()
    sharepoint_drive_id = settings.sharepoint_drive_id
    base_folder_path = settings.base_folder_path or "General Management/Công văn"

    try:
        frappe.msgprint("Starting SharePoint upload process...")
        frappe.msgprint(f"Fetching File Doc: {file_doc_name}")
        file_doc = frappe.get_doc("File", file_doc_name)
        frappe.msgprint(f"Got File Doc. Path: {file_doc.file_url}")
        file_path_rel = file_doc.get_full_path().lstrip('/')
        file_path_abs = get_site_path(file_path_rel)
        if not os.path.exists(file_path_abs):
            frappe.throw(f"File not found at path: {file_path_abs}")

        file_name = file_doc.file_name

        # --- Determine Target Folder ---
        year = frappe.utils.now_datetime().strftime("%Y")
        doctype_folder = doc.doctype.replace(" ", "_")
        safe_doc_name = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in doc.name).rstrip()
        doc_name_folder = safe_doc_name

        target_folder_rel_path = f"{base_folder_path}/{year}/{doctype_folder}/{doc_name_folder}"
        target_folder_rel_path = target_folder_rel_path.strip('/')

        # --- Ensure Folder Exists ---
        frappe.msgprint(f"Ensuring SharePoint folder exists: '{target_folder_rel_path}'")
        folder_id = create_sharepoint_folder_if_not_exists(sharepoint_drive_id, target_folder_rel_path)
        if not folder_id:
             # Error handled within create_sharepoint_folder_if_not_exists
             return None

        # --- Upload File ---
        frappe.msgprint(f"Target SharePoint folder confirmed (ID: {folder_id}). Preparing file upload...")
        encoded_file_name = encode(file_name)
        # Use item ID for parent folder reference - more reliable than path
        upload_url_base = f"https://graph.microsoft.com/v1.0/drives/{sharepoint_drive_id}/items/{folder_id}:/{encoded_file_name}:"

        file_size = os.path.getsize(file_path_abs)
        if file_size > 4 * 1024 * 1024:
             # TODO: Implement resumable upload using createUploadSession
             # upload_url = f"{upload_url_base}/createUploadSession"
             # ... implementation needed ...
             frappe.throw("File size exceeds 4MB. Resumable upload not yet implemented.")
             # upload_result = upload_large_file(...)
        else:
            # Simple PUT for smaller files
            upload_url = f"{upload_url_base}/content"
            headers = get_headers()
            headers["Content-Type"] = file_doc.content_type or "application/octet-stream"

            with open(file_path_abs, "rb") as f:
                file_content = f.read()

            # frappe.log_info(f"Attempting SharePoint small file upload to item ID '{folder_id}' with name '{encoded_file_name}'")
            frappe.msgprint(f"Uploading file '{file_name}' to SharePoint folder '{target_folder_rel_path}'...")
            response = requests.put(upload_url, headers=headers, data=file_content)
            response.raise_for_status() # Raise HTTPError for bad responses
            upload_result = response.json()
            frappe.msgprint(f"File '{file_name}' uploaded successfully.")

        frappe.log_info(f"SharePoint upload response: {upload_result}")

        # --- Create Document Version Entry ---
        frappe.msgprint("Creating document version entry...")
        sharepoint_link = upload_result.get("webUrl")
        # Use the item ID from the upload response as a more stable version indicator if available
        version_id = upload_result.get("id", upload_result.get("eTag", "N/A"))

        current_version_count = len(doc.get("versions", []))
        next_version_number = current_version_count + 1

        new_version = doc.append("versions", {
            "version_number": next_version_number,
            "sharepoint_link": sharepoint_link,
            "action_taken": action_details.get("action", "Upload"),
            "action_by": action_details.get("user", frappe.session.user),
            "action_timestamp": frappe.utils.now_datetime(),
            "file_url": file_doc.file_url if doc.get("store_locally") else None
        })

        # --- Optional: Delete Local File ---
        if not doc.get("store_locally"):
            try:
                os.remove(file_path_abs)
                # Optionally detach or delete the Frappe File record
                # file_doc.db_set("attached_to_doctype", None)
                # file_doc.db_set("attached_to_name", None)
                # frappe.delete_doc("File", file_doc.name, ignore_permissions=True, force=True) # Be very careful with force=True
                # frappe.log_info(f"Removed local file '{file_path_abs}' after SharePoint upload for '{doc.name}'.")
                frappe.msgprint(f"Removed local file '{file_path_abs}'.")
                # Update the child table entry's file_url after deletion
                frappe.db.set_value("Document Version", new_version.name, "file_url", None)
            except Exception as del_err:
                # Log deletion error but don't necessarily stop the whole process, maybe just msgprint?
                # Or throw if deletion is critical? User asked for throw.
                frappe.throw(f"Failed to delete local file '{file_path_abs}': {del_err}")


        # frappe.log_info(f"Successfully uploaded '{file_name}' to SharePoint for document '{doc.name}'. Link: {sharepoint_link}")
        frappe.msgprint(f"Document version created. SharePoint Link: {sharepoint_link}")
        # Save the document to persist the new child table row
        doc.save(ignore_permissions=True) # Save needed to persist child table changes made via .append()
        frappe.msgprint("SharePoint upload process completed successfully.")
        return {"sharepoint_link": sharepoint_link, "version_id": version_id}

    except requests.exceptions.RequestException as e:
        err_msg = e.response.text if e.response else str(e)
        # Removed redundant log_error before throw
        frappe.throw(f"Failed to upload file to SharePoint: {err_msg}")
        return None
    except Exception as e:
        # Removed redundant log_error before throw
        frappe.throw(f"An unexpected error occurred during SharePoint upload: {e}")
        return None

# TODO: Implement upload_large_file function using createUploadSession
# def upload_large_file(drive_id, parent_folder_id, file_name, file_path_abs):
#     pass

# TODO: Add functions for other interactions if needed (e.g., get_file, delete_file, create_folder)