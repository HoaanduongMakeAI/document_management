
# For license information, please see license.txt

import frappe
import requests
import os
import urllib.parse
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

def get_group_details_and_sharepoint_ids(group_id):
    """
    Fetches the Microsoft Entra Group Name, SharePoint Site ID, and the default Document Library Drive ID for a given Microsoft Entra Group ID.

    Args:
        group_id (str): The ID of the Microsoft Entra Group.

    Returns:
        tuple: (group_name, site_id, drive_id) or raises an exception on failure.
    """
    if not group_id:
        frappe.throw("Microsoft Entra Group ID is required to fetch details.")

    # Ensure Connected App is configured (needed for get_headers)
    settings = get_sharepoint_settings() # This checks for connected_app implicitly

    headers = get_headers() # Fetches token using settings.connected_app

    try:
        # Get Group Details (including name)
        group_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}"
        frappe.msgprint(f"Fetching group details from: {group_url}")
        group_response = requests.get(group_url, headers=headers)
        group_response.raise_for_status() # Check for HTTP errors
        group_data = group_response.json()
        group_name = group_data.get("displayName")
        if not group_name:
             # Log a warning if display name is missing but don't necessarily fail
             frappe.log_warning(f"Group display name not found for Group ID: {group_id}. Response: {group_data}", "SharePoint Integration Warning")
             group_name = "N/A" # Set a default or indicate it was not found

        frappe.msgprint(f"Found Group Name: {group_name}")


        # Get SharePoint site associated with the group
        site_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/sites/root"
        frappe.msgprint(f"Fetching site info from: {site_url}")
        site_response = requests.get(site_url, headers=headers)
        site_response.raise_for_status() # Check for HTTP errors
        site_data = site_response.json()
        site_id = site_data.get("id")
        if not site_id:
            frappe.throw(f"Could not retrieve SharePoint Site ID for Group ID: {group_id}. Response: {site_data}")
        frappe.msgprint(f"Found Site ID: {site_id}")

        # Get drives (document libraries) within the site
        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        frappe.msgprint(f"Fetching drives from: {drives_url}")
        drives_response = requests.get(drives_url, headers=headers)
        drives_response.raise_for_status()
        drives_data = drives_response.json()
        drive_id = None
        if drives_data.get("value"):
            # Prioritize the 'Documents' library if it exists
            for drive in drives_data["value"]:
                if drive.get("name", "").lower() == "documents":
                    drive_id = drive.get("id")
                    frappe.msgprint(f"Found 'Documents' Drive ID: {drive_id}")
                    break
            # Fallback to the first drive if 'Documents' is not found
            if not drive_id and drives_data["value"]:
                drive_id = drives_data["value"][0].get("id")
                frappe.msgprint(f"Using first Drive ID as fallback: {drive_id}")
        
        if not drive_id:
            frappe.throw(f"Could not retrieve any Drive ID for Site ID: {site_id}. Response: {drives_data}")

        return group_name, site_id, drive_id

    except requests.exceptions.RequestException as e:
        err_msg = e.response.text if e.response else str(e)
        frappe.log_error(f"Graph API Error fetching details for group {group_id}: {err_msg}", "SharePoint Integration Error")
        frappe.throw(f"Error communicating with Microsoft Graph API while fetching details for group {group_id}: {err_msg}")
    except Exception as e:
        frappe.log_error(f"Unexpected error fetching details for group {group_id}: {frappe.get_traceback()}", "SharePoint Integration Error")
        frappe.throw(f"An unexpected error occurred while fetching SharePoint details for group {group_id}: {str(e)}")
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
        # Use the raw segment name for filtering. Requests library handles basic URL encoding.
        # Avoid frappe.utils.encode here as it seems to cause issues with the filter format.
        segment_name_for_filter = segment

        # Graph API URL to get/create a folder within another item (parent)
        # Using item-id based addressing is more robust than path-based for creation
        # Use requests params for proper URL encoding
        folder_check_base_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent_item_id}/children"
        # Use the raw segment string in the filter, ensuring it's properly quoted for OData
        params = {"$filter": f"name eq '{segment_name_for_filter}'"}

        try:
            frappe.msgprint(f"Checking for SharePoint folder segment: '{segment}'...")
            # Pass params dict to requests.get for automatic encoding
            response = requests.get(folder_check_base_url, headers=headers, params=params)
            response.raise_for_status()
            data = response.json()

            if data.get("value"): # Folder exists
                parent_item_id = data["value"][0]["id"]
                frappe.msgprint(f"SharePoint folder segment '{segment}' exists.")
                # frappe.msgprint(f"SharePoint folder segment '{segment}' exists with ID: {parent_item_id}") # Keep log for detailed ID if needed
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
                # frappe.msgprint(f"Created SharePoint folder segment '{segment}' with ID: {parent_item_id}") # Keep log for detailed ID if needed

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
    try:
        frappe.msgprint("Starting SharePoint upload process...")

        # --- Determine Target Drive and Folder Path ---
        target_folder_rel_path = None
        sharepoint_drive_id = None

        if doc.get("folder"):
            try:
                folder_doc = frappe.get_doc("Folder", doc.folder)
                if not folder_doc.microsoft_entra_group:
                     frappe.throw(f"Folder '{doc.folder}' does not have a Microsoft Entra Group linked.")
                
                group_doc = frappe.get_doc("Microsoft Entra Group", folder_doc.microsoft_entra_group)
                if not group_doc.sharepoint_drive_id:
                     frappe.throw(f"Linked Microsoft Entra Group '{folder_doc.microsoft_entra_group}' for Folder '{doc.folder}' is missing its SharePoint Drive ID.")

                sharepoint_drive_id = group_doc.sharepoint_drive_id
                # Use the path directly from the Folder doc. Do not strip slashes here.
                target_folder_rel_path = folder_doc.folder_path
                frappe.msgprint(f"Using specified Folder: '{doc.folder}' (Path: {target_folder_rel_path}, Drive: {sharepoint_drive_id})")

            except frappe.DoesNotExistError:
                 frappe.throw(f"Specified Folder '{doc.folder}' not found.")
            except Exception as e:
                 frappe.throw(f"Error fetching details for specified Folder '{doc.folder}': {e}")
        else:
            frappe.msgprint("No specific Folder selected, using default settings.")
            settings = get_sharepoint_settings() # Fetches settings, validates required fields like connected_app, entra_group_id
            if not settings.sharepoint_drive_id:
                 # Attempt to fetch IDs if missing in settings
                 # Corrected import path below
                 from document_management.document_management.doctype.document_management_settings.document_management_settings import fetch_sharepoint_ids_from_group
                 try:
                     ids = fetch_sharepoint_ids_from_group()
                     settings.reload() # Reload to get updated IDs
                     sharepoint_drive_id = settings.sharepoint_drive_id
                     if not sharepoint_drive_id:
                          frappe.throw("SharePoint Drive ID is still missing in Document Management Settings even after attempting to fetch.")
                 except Exception as fetch_e:
                      frappe.throw(f"SharePoint Drive ID is missing in Document Management Settings and failed to fetch automatically: {fetch_e}")
            else:
                 sharepoint_drive_id = settings.sharepoint_drive_id

            # Use base_folder_path from settings, default if empty. Do not strip slashes here.
            target_folder_rel_path = settings.base_folder_path or "Uncategorized"
            frappe.msgprint(f"Using default settings path: '{target_folder_rel_path}' (Drive: {sharepoint_drive_id})")

        if not sharepoint_drive_id or target_folder_rel_path is None:
             frappe.throw("Could not determine SharePoint Drive ID or target folder path.")

        # Validate the determined path is not empty or just root after potential whitespace strip
        if not target_folder_rel_path.strip() or target_folder_rel_path.strip() == "/":
             frappe.throw(f"Target folder path '{target_folder_rel_path}' is invalid.")

        frappe.msgprint(f"Fetching File Doc: {file_doc_name}")
        file_doc = frappe.get_doc("File", file_doc_name)
        frappe.msgprint(f"Got File Doc. Path: {file_doc.file_url}")
        # Use get_files_path for more robust file path resolution
        file_path_abs = get_files_path(file_doc.file_name, is_private=file_doc.is_private)
        if not os.path.exists(file_path_abs):
            # Log the expected path for debugging if it still fails
            frappe.log_error(f"Expected file path: {file_path_abs}", "SharePoint Integration File Not Found")
            frappe.throw(f"File not found at path: {file_path_abs}")

        file_name = file_doc.file_name

        # --- Ensure Base Folder Exists ---
        # The target_folder_rel_path is now determined above based on doc.folder or settings.
        # Pass the raw path to the function, it handles stripping/splitting.
        frappe.msgprint(f"Ensuring SharePoint base folder exists: '{target_folder_rel_path}'")
        folder_id = create_sharepoint_folder_if_not_exists(sharepoint_drive_id, target_folder_rel_path)
        if not folder_id:
             # Error handled within create_sharepoint_folder_if_not_exists
             return None

        # --- Upload File ---
        frappe.msgprint(f"Target SharePoint folder confirmed (ID: {folder_id}). Preparing file upload...")
        encoded_file_name = urllib.parse.quote(file_name, safe='')
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
            # Safely get content_type, default if missing
            content_type = getattr(file_doc, "content_type", None) or "application/octet-stream"
            headers["Content-Type"] = content_type

            with open(file_path_abs, "rb") as f:
                file_content = f.read()

            # frappe.msgprint(f"Attempting SharePoint small file upload to item ID '{folder_id}' with name '{encoded_file_name}'")
            # The file is uploaded *into* the folder represented by folder_id.
            # The target_folder_rel_path variable holds the logical path used to get the folder_id.
            frappe.msgprint(f"Uploading file '{file_name}' into SharePoint folder (Path used: '{target_folder_rel_path}', Target Item ID: '{folder_id}')...")
            response = requests.put(upload_url, headers=headers, data=file_content)
            response.raise_for_status() # Raise HTTPError for bad responses
            upload_result = response.json()
            frappe.msgprint(f"File '{file_name}' uploaded successfully.")

        frappe.msgprint(f"SharePoint upload response: {upload_result}")

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
                # frappe.msgprint(f"Removed local file '{file_path_abs}' after SharePoint upload for '{doc.name}'.")
                frappe.msgprint(f"Removed local file '{file_path_abs}'.")
                # Update the child table entry's file_url after deletion
                frappe.db.set_value("Document Version", new_version.name, "file_url", None)
            except Exception as del_err:
                # Log deletion error but don't necessarily stop the whole process, maybe just msgprint?
                # Or throw if deletion is critical? User asked for throw.
                frappe.throw(f"Failed to delete local file '{file_path_abs}': {del_err}")


        # frappe.msgprint(f"Successfully uploaded '{file_name}' to SharePoint for document '{doc.name}'. Link: {sharepoint_link}")
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

@frappe.whitelist()
def list_sharepoint_folder_contents(folder_docname, relative_path="/"):
    """
    Lists files and folders within a given relative path of a SharePoint folder
    defined by a 'Folder' DocType.

    :param folder_docname: The name of the 'Folder' DocType record.
    :param relative_path: The relative path within the SharePoint folder (e.g., "/", "/SubfolderA").
    :return: dict with "items": list of {"name": str, "path": str, "is_folder": bool} or {"error": str}
    """
    if not folder_docname:
        return {"error": "Folder DocType name is required."}

    try:
        folder_doc = frappe.get_doc("Folder", folder_docname)
        if not folder_doc.microsoft_entra_group:
            return {"error": f"Folder '{folder_docname}' does not have a Microsoft Entra Group linked."}
        
        group_doc = frappe.get_doc("Microsoft Entra Group", folder_doc.microsoft_entra_group)
        if not group_doc.sharepoint_drive_id:
            return {"error": f"Linked Microsoft Entra Group '{folder_doc.microsoft_entra_group}' for Folder '{folder_docname}' is missing its SharePoint Drive ID."}

        drive_id = group_doc.sharepoint_drive_id
        # Base path from Folder doctype, ensure it's clean and doesn't start/end with / for joining
        base_sp_path_from_folder_doc = (folder_doc.folder_path or "").strip("/")

        # Clean the relative_path input
        clean_relative_path = relative_path.strip("/")

        # Construct the full path for the API call
        # If base_sp_path_from_folder_doc is empty, path_for_api is just clean_relative_path
        # If clean_relative_path is empty (i.e. root of relative_path), path_for_api is base_sp_path_from_folder_doc
        # Otherwise, join them.
        if base_sp_path_from_folder_doc and clean_relative_path:
            path_for_api = f"{base_sp_path_from_folder_doc}/{clean_relative_path}"
        elif base_sp_path_from_folder_doc:
            path_for_api = base_sp_path_from_folder_doc
        elif clean_relative_path:
            path_for_api = clean_relative_path
        else: # Both are empty or "/", target the root of the drive
            path_for_api = ""


        headers = get_headers()
        items = []

        if not path_for_api: # Root of the drive
            graph_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
        else:
            encoded_path = urllib.parse.quote(path_for_api)
            graph_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_path}:/children"

        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        children = response.json().get("value", [])

        for item in children:
            item_name = item.get("name")
            # The path we return should be relative to the Folder doctype's root,
            # so it's essentially the 'clean_relative_path' + item_name
            item_relative_path = f"/{clean_relative_path}/{item_name}".replace("//", "/") if clean_relative_path else f"/{item_name}"
            
            items.append({
                "name": item_name,
                "path": item_relative_path, # This path is relative to the Folder DocType's root + initial relative_path
                "is_folder": "folder" in item
            })
        return {"items": items}

    except requests.exceptions.RequestException as e:
        err_msg = e.response.text if e.response else str(e)
        frappe.log_error(f"Graph API error listing contents for {folder_docname} at '{relative_path}': {err_msg}", "SharePoint List Contents Error")
        return {"error": f"API Error: {err_msg}"}
    except frappe.DoesNotExistError as e:
        frappe.log_error(f"DoesNotExistError for {folder_docname}: {e}", "SharePoint List Contents Error")
        return {"error": str(e)}
    except Exception as e:
        frappe.log_error(f"Error listing SharePoint contents for {folder_docname} at '{relative_path}': {frappe.get_traceback()}", "SharePoint List Contents Error")
        return {"error": f"Unexpected error: {str(e)}"}


@frappe.whitelist()
def get_sharepoint_item_details(target_folder_docname, relative_path_to_item):
    """
    Gets details (link, full path, type) of a specific item in SharePoint.
    The relative_path_to_item is relative to the target_folder_docname's root.
    """
    if not target_folder_docname or not relative_path_to_item:
        return {"error": "Target Folder DocName and relative path to item are required."}

    try:
        folder_doc = frappe.get_doc("Folder", target_folder_docname)
        if not folder_doc.microsoft_entra_group:
            return {"error": f"Folder '{target_folder_docname}' does not have a Microsoft Entra Group linked."}
        
        group_doc = frappe.get_doc("Microsoft Entra Group", folder_doc.microsoft_entra_group)
        if not group_doc.sharepoint_drive_id:
            return {"error": f"Linked Microsoft Entra Group for Folder '{target_folder_docname}' is missing its SharePoint Drive ID."}

        drive_id = group_doc.sharepoint_drive_id
        base_sp_path_from_folder_doc = (folder_doc.folder_path or "").strip("/")
        
        # relative_path_to_item is already relative to the Folder DocType's root
        clean_relative_path_to_item = relative_path_to_item.strip("/")

        if base_sp_path_from_folder_doc and clean_relative_path_to_item:
            full_item_path_for_api = f"{base_sp_path_from_folder_doc}/{clean_relative_path_to_item}"
        elif base_sp_path_from_folder_doc: # Should not happen if relative_path_to_item is to a specific item
            full_item_path_for_api = base_sp_path_from_folder_doc
        elif clean_relative_path_to_item:
            full_item_path_for_api = clean_relative_path_to_item
        else:
            return {"error": "Invalid item path provided."}


        headers = get_headers()
        encoded_full_item_path = urllib.parse.quote(full_item_path_for_api)
        graph_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_full_item_path}"
        
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        item_data = response.json()

        is_file = "file" in item_data
        # The 'path' in parentReference might be useful for constructing absolute_path if needed differently
        # For now, use the known full_item_path_for_api as the basis for absolute_path
        # Construct a user-friendly absolute path, e.g., "Shared Documents/FolderA/File.txt"
        # The item_data.get("parentReference", {}).get("path") gives path from drive root.
        # Example: /drives/b!..../root:/FolderFromDoc/RelPathToItem
        # We need to extract the part after "root:"
        sp_path_from_drive_root = item_data.get("parentReference", {}).get("path", "")
        if sp_path_from_drive_root and 'root:' in sp_path_from_drive_root:
            sp_path_from_drive_root = sp_path_from_drive_root.split('root:', 1)[1]
        
        # Ensure it starts with a slash if not empty
        if sp_path_from_drive_root and not sp_path_from_drive_root.startswith('/'):
             sp_path_from_drive_root = '/' + sp_path_from_drive_root

        absolute_path = f"{sp_path_from_drive_root}/{item_data.get('name')}".replace("//","/")


        return {
            "sharepoint_link": item_data.get("webUrl"),
            "absolute_path": absolute_path,
            "name": item_data.get("name"),
            "is_file": is_file,
            "id": item_data.get("id")
        }

    except requests.exceptions.RequestException as e:
        err_msg = e.response.text if e.response else str(e)
        frappe.log_error(f"Graph API error getting item details for {target_folder_docname} at '{relative_path_to_item}': {err_msg}", "SharePoint Item Details Error")
        return {"error": f"API Error: {err_msg}"}
    except frappe.DoesNotExistError as e:
        frappe.log_error(f"DoesNotExistError for {target_folder_docname}: {e}", "SharePoint Item Details Error")
        return {"error": str(e)}
    except Exception as e:
        frappe.log_error(f"Error getting SharePoint item details for {target_folder_docname} at '{relative_path_to_item}': {frappe.get_traceback()}", "SharePoint Item Details Error")
        return {"error": f"Unexpected error: {str(e)}"}


@frappe.whitelist()
def upload_file_to_path(doctype, docname, file_doc_name, target_folder_docname, target_relative_path="/"):
    """
    Uploads a file to a specific relative path within a SharePoint folder defined by 'Folder' DocType.
    Updates the source document's 'teams_link' and 'path' fields.
    target_relative_path is relative to the target_folder_docname's root.
    """
    if not doctype or not docname or not file_doc_name or not target_folder_docname:
        return {"error": "Doctype, Docname, File Doc Name, and Target Folder DocName are required."}

    try:
        doc = frappe.get_doc(doctype, docname)
        file_doc = frappe.get_doc("File", file_doc_name)

        folder_settings_doc = frappe.get_doc("Folder", target_folder_docname)
        if not folder_settings_doc.microsoft_entra_group:
            return {"error": f"Target Folder '{target_folder_docname}' does not have a Microsoft Entra Group linked."}

        group_doc = frappe.get_doc("Microsoft Entra Group", folder_settings_doc.microsoft_entra_group)
        if not group_doc.sharepoint_drive_id:
            return {"error": f"Linked Microsoft Entra Group for Target Folder '{target_folder_docname}' is missing its SharePoint Drive ID."}

        drive_id = group_doc.sharepoint_drive_id
        # Base path from the 'Folder' doctype record.
        base_sp_path_from_folder_doc = (folder_settings_doc.folder_path or "").strip("/")
        
        # target_relative_path is the path *within* the Folder DocType's location.
        clean_target_relative_path = target_relative_path.strip("/")

        # Construct the final SharePoint folder path for upload by combining base and relative.
        if base_sp_path_from_folder_doc and clean_target_relative_path:
            final_sp_folder_path = f"{base_sp_path_from_folder_doc}/{clean_target_relative_path}"
        elif base_sp_path_from_folder_doc:
            final_sp_folder_path = base_sp_path_from_folder_doc
        elif clean_target_relative_path:
            final_sp_folder_path = clean_target_relative_path
        else: # Uploading to the root of the folder_settings_doc.folder_path or drive root if base is also empty
            final_sp_folder_path = "" # Represents root for create_sharepoint_folder_if_not_exists logic

        # Ensure the target folder structure exists in SharePoint
        # If final_sp_folder_path is empty, create_sharepoint_folder_if_not_exists should correctly use "root" for parent_item_id.
        target_sp_folder_id = create_sharepoint_folder_if_not_exists(drive_id, final_sp_folder_path if final_sp_folder_path else "/")
        if not target_sp_folder_id:
            return {"error": "Failed to ensure target SharePoint folder exists."} # Error already thrown by create_sharepoint_folder_if_not_exists

        file_path_abs = get_files_path(file_doc.file_name, is_private=file_doc.is_private)
        if not os.path.exists(file_path_abs):
            return {"error": f"Local file not found at path: {file_path_abs}"}

        file_name_on_sp = file_doc.file_name
        encoded_file_name_on_sp = urllib.parse.quote(file_name_on_sp, safe='')
        
        # Upload URL using the ID of the target folder
        upload_url_base = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{target_sp_folder_id}:/{encoded_file_name_on_sp}:"
        
        file_size = os.path.getsize(file_path_abs)
        upload_api_response_data = None

        if file_size > 4 * 1024 * 1024: # More than 4MB
            # TODO: Implement resumable upload for upload_file_to_path
            frappe.throw("File size exceeds 4MB. Resumable upload for specific path not yet implemented in this function.")
        else:
            upload_url = f"{upload_url_base}/content"
            headers = get_headers()
            content_type = getattr(file_doc, "content_type", None) or "application/octet-stream"
            headers["Content-Type"] = content_type
            with open(file_path_abs, "rb") as f:
                file_content = f.read()
            
            response = requests.put(upload_url, headers=headers, data=file_content)
            response.raise_for_status()
            upload_api_response_data = response.json()

        if not upload_api_response_data or not upload_api_response_data.get("webUrl"):
            return {"error": "Upload to SharePoint succeeded but no webUrl returned."}

        sharepoint_link = upload_api_response_data.get("webUrl")
        
        # Construct user-friendly absolute path
        sp_path_from_drive_root_for_file = upload_api_response_data.get("parentReference", {}).get("path", "")
        if sp_path_from_drive_root_for_file and 'root:' in sp_path_from_drive_root_for_file:
            sp_path_from_drive_root_for_file = sp_path_from_drive_root_for_file.split('root:', 1)[1]
        if sp_path_from_drive_root_for_file and not sp_path_from_drive_root_for_file.startswith('/'):
             sp_path_from_drive_root_for_file = '/' + sp_path_from_drive_root_for_file
        
        absolute_file_path_on_sp = f"{sp_path_from_drive_root_for_file}/{file_name_on_sp}".replace("//","/")

        # Update the original document
        doc.set("teams_link", sharepoint_link)
        doc.set("path", absolute_file_path_on_sp) # New field
        doc.set("folder", target_folder_docname) # Set the folder field to the one used for upload
        
        # Create a version entry (similar to upload_file_to_sharepoint)
        action_details = { "action": "Upload to Path", "user": frappe.session.user }
        current_version_count = len(doc.get("versions", []))
        next_version_number = current_version_count + 1
        new_version = doc.append("versions", {
            "version_number": next_version_number,
            "sharepoint_link": sharepoint_link,
            "action_taken": action_details.get("action"),
            "action_by": action_details.get("user"),
            "action_timestamp": frappe.utils.now_datetime(),
            "file_url": file_doc.file_url if doc.get("store_locally") else None,
            "path": absolute_file_path_on_sp # Store path in version too
        })
        doc.save(ignore_permissions=True) # Save to persist changes and new version

        if not doc.get("store_locally"):
            try:
                os.remove(file_path_abs)
                frappe.db.set_value("Document Version", new_version.name, "file_url", None)
            except Exception as del_err:
                frappe.log_error(f"Failed to delete local file '{file_path_abs}' after SP upload: {del_err}", "SharePoint Upload to Path")
                # Non-critical, so don't throw, just log.

        return {
            "sharepoint_link": sharepoint_link,
            "absolute_path": absolute_file_path_on_sp,
            "version_id": upload_api_response_data.get("id", upload_api_response_data.get("eTag"))
        }

    except requests.exceptions.RequestException as e:
        err_msg = e.response.text if e.response else str(e)
        frappe.log_error(f"Graph API error uploading to path: {err_msg}", "SharePoint Upload to Path Error")
        return {"error": f"API Error: {err_msg}"}
    except frappe.DoesNotExistError as e:
        frappe.log_error(f"DoesNotExistError during upload to path: {e}", "SharePoint Upload to Path Error")
        return {"error": str(e)}
    except Exception as e:
        frappe.log_error(f"Error uploading to SharePoint path: {frappe.get_traceback()}", "SharePoint Upload to Path Error")
        return {"error": f"Unexpected error: {str(e)}"}


# TODO: Implement upload_large_file function using createUploadSession
# def upload_large_file(drive_id, parent_folder_id, file_name, file_path_abs):
#     pass

# TODO: Add functions for other interactions if needed (e.g., get_file, delete_file, create_folder)