# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from frappe.utils import strip_html, get_abbr

# Import the upload function - Adjust path due to directory move
from document_management.document_management.utils.sharepoint_integration import upload_file_to_sharepoint, get_sharepoint_version_from_link

class IncomingDocument(Document):
	# autoname is now handled by Naming Series in JSON
	# before_save hook is removed. Upload logic is now triggered manually via upload_file_via_modal.

	def before_save(self):
		# Store original status before save
		if self.name: # Check if it's an existing document
			self._original_status = frappe.db.get_value("Incoming Document", self.name, "status")
			self._original_team_links = frappe.db.get_value("Incoming Document", self.name, "team_links") # Store original team_links
			# Fetch and store original document tasks as a dictionary
			original_tasks = frappe.get_all("Document Task", filters={"parent": self.name}, fields=["*"])
			self._original_tasks_dict = {d.name: d for d in original_tasks} # Create dictionary for easy lookup
			frappe.msgprint(f"Original tasks dict (before_save): {self._original_tasks_dict}")
		else: # New document
			self._original_status = None
			self._original_tasks_dict = {}
			self._original_team_links = None # Initialize for new document

	def on_update(self):
		# Check for changes in team_links
		if self.team_links != (self._original_team_links if hasattr(self, '_original_team_links') else None):
			if self.team_links: # Only trigger if team_links is not empty after change
				self.create_document_version_and_notify("Teams Link Updated")

		# Check status changes to trigger notifications
		if self.status != self._original_status:
			if self.status == "Under Review":
				self.notify_reviewers()
			elif self.status == "Tasks Assigned":
				self.notify_assigned_users()

		# Check for changes in document tasks and group by assignee
		original_tasks_dict = self._original_tasks_dict if hasattr(self, '_original_tasks_dict') and self._original_tasks_dict else {}
		current_tasks_dict = {d.name: d.as_dict() for d in self.document_tasks}

		frappe.msgprint(f"Original tasks dict (on_update): {original_tasks_dict}")
		frappe.msgprint(f"Current tasks dict (on_update): {current_tasks_dict}")

		tasks_to_notify = {} # {assignee: [{task_details, change_type}]}

		# Check for new or modified tasks
		for task_name, current_task in current_tasks_dict.items():
			original_task = original_tasks_dict.get(task_name)
			change_type = None

			if original_task is None:
				# New task
				change_type = "New"
			else:
				# Existing task, check for modifications
				# Compare relevant fields: content, assignee, task_status, due_date
				if (current_task.get("content") != original_task.get("content") or
					current_task.get("assignee") != original_task.get("assignee") or
					current_task.get("task_status") != original_task.get("task_status") or
					str(current_task.get("due_date")) != str(original_task.get("due_date"))): # Compare dates as strings
					change_type = "Updated"

			if change_type and current_task.get("assignee"):
				if current_task["assignee"] not in tasks_to_notify:
					tasks_to_notify[current_task["assignee"]] = []
				tasks_to_notify[current_task["assignee"]].append({"task": frappe._dict(current_task), "change_type": change_type})
		print(tasks_to_notify)

		# Send consolidated email to each assignee
		for assignee, tasks in tasks_to_notify.items():
			self.notify_assignee_tasks_change(assignee, tasks)

	def notify_assignee_tasks_change(self, assignee, tasks):
		if not tasks:
			return

		subject = f"Cập nhật công việc liên quan đến Văn bản đến: {self.incoming_number} - {self.subject}"
		body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Có cập nhật công việc liên quan đến văn bản đến "{self.incoming_number} - {self.subject}":</p>
<ul>
"""
		for item in tasks:
			task = item["task"]
			change_type = item["change_type"]
			body += f"""
	<li>
		<strong>Loại thay đổi:</strong> {change_type}<br>
		<strong>Nội dung:</strong> {task.content}<br>
		<strong>Trạng thái:</strong> {task.task_status}<br>
		<strong>Deadline:</strong> {task.due_date or 'N/A'}
	</li>
"""
		body += f"""
</ul>
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết văn bản và các công việc:</p>
<p><a href="/app/incoming-document/{self.name}">Xem Văn bản đến trên ERPNext</a></p>
"""
		if self.teams_link:
			body += f"""
<p>File văn bản gốc có thể xem tại đây:</p>
<p><a href="{self.teams_link}">Xem File trên Microsoft Teams</a></p>
"""

		body += """
<p>Trân trọng,</p>
<p>Hệ thống ERPNext</p>
"""

		try:
			frappe.sendmail(
				recipients=[assignee],
				subject=subject,
				message=body,
				now=True # Send immediately
			)
			frappe.log_error(f"Consolidated task change notification sent for Incoming Document: {self.name} to {assignee}", "INCOMING DOCUMENT CONSOLIDATED TASK NOTIFICATION SENT")
		except Exception as e:
			frappe.log_error(f"Failed to send consolidated task change notification for Incoming Document: {self.name} to {assignee}: {e}", f"INCOMING DOCUMENT CONSOLIDATED TASK NOTIFICATION FAILED")

	# Keep notify_reviewers method as it contains email logic for reviewers

	def notify_reviewers(self):
		# Get users with 'System Manager' or 'Document Manager' roles
		reviewers = frappe.get_all("User",
			filters={
				"user_type": "System User",
				"enabled": 1
			},
			or_filters=[
				["User role", "role", "=", "System Manager"],
				["User role", "role", "=", "Document Manager"],
				["User role", "role", "=", "Văn thư"]
			],
			pluck="email"
		)

		if not reviewers:
			frappe.log_error(f"No reviewers found for Incoming Document: {self.name}", "INCOMING DOCUMENT REVIEW NOTIFICATION FAILED")
			return

		# Construct email subject and body
		subject = f"Yêu cầu xử lý Văn bản đến: {self.incoming_number} - {self.subject}"
		body = f"""
<p>Kính gửi Ban Lãnh đạo/Cố vấn,</p>
<p>Có một văn bản đến mới cần được xem xét và chỉ đạo xử lý:</p>
<ul>
	<li><strong>Số đến:</strong> {self.incoming_number}</li>
	<li><strong>Trích yếu:</strong> {self.subject}</li>
	<li><strong>Nơi gửi:</strong> {self.sender}</li>
	<li><strong>Ngày đến:</strong> {self.date_received}</li>
</ul>
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết và đưa ra ý kiến chỉ đạo:</p>
<p><a href="/app/incoming-document/{self.name}">Xem Văn bản đến trên ERPNext</a></p>
"""

		if self.teams_link:
			body += f"""
<p>File văn bản gốc có thể xem tại đây:</p>
<p><a href="{self.teams_link}">Xem File trên Microsoft Teams</a></p>
"""

		body += """
<p>Trân trọng,</p>
<p>Hệ thống ERPNext</p>
"""

		# Send email
		frappe.sendmail(
			recipients=reviewers,
			subject=subject,
			message=body,
			now=True # Send immediately
		)

		frappe.log_error(f"Notification sent for Incoming Document: {self.name} to {', '.join(reviewers)}", "INCOMING DOCUMENT REVIEW NOTIFICATION SENT")


	def notify_assigned_users(self):
		# Get users assigned to tasks in the Document Task child table
		assigned_users = []
		if self.document_tasks:
			for task in self.document_tasks:
				if task.assignee:
					assigned_users.append(task.assignee)

		# Remove duplicates and current user from the list
		assigned_users = list(set(assigned_users))
		if frappe.session.user in assigned_users:
			assigned_users.remove(frappe.session.user)

		if not assigned_users:
			frappe.log_error(f"No users assigned to tasks for Incoming Document: {self.name}", "INCOMING DOCUMENT ASSIGNMENT NOTIFICATION FAILED")
			return

		# Construct email subject and body
		subject = f"Bạn có công việc mới liên quan đến Văn bản đến: {self.incoming_number} - {self.subject}"
		body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Bạn được giao xử lý văn bản đến sau:</p>
<ul>
	<li><strong>Số đến:</strong> {self.incoming_number}</li>
	<li><strong>Trích yếu:</strong> {self.subject}</li>
	<li><strong>Nơi gửi:</strong> {self.sender}</li>
	<li><strong>Ngày đến:</strong> {self.date_received}</li>
</ul>
"""

		if self.instructions:
			body += f"""
<p><strong>Ý kiến chỉ đạo chung:</strong></p>
<p>{self.instructions}</p>
"""

		if self.document_tasks:
			body += """
<p><strong>Chi tiết công việc được giao:</strong></p>
<ul>
"""
			for task in self.document_tasks:
				if task.assignee in assigned_users:
					body += f"""
	<li>
		<strong>Nội dung:</strong> {task.content}<br>
		<strong>Trạng thái:</strong> {task.task_status}<br>
		<strong>Deadline:</strong> {task.due_date or 'N/A'}
	</li>
"""
			body += "</ul>"


		body += f"""
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết văn bản và các công việc:</p>
<p><a href="/app/incoming-document/{self.name}">Xem Văn bản đến trên ERPNext</a></p>
"""

		if self.teams_link:
			body += f"""
<p>File văn bản gốc có thể xem tại đây:</p>
<p><a href="{self.teams_link}">Xem File trên Microsoft Teams</a></p>
"""

		body += """
<p>Trân trọng,</p>
<p>Hệ thống ERPNext</p>
"""

		# Send email
		frappe.sendmail(
			recipients=assigned_users,
			subject=subject,
			message=body,
			now=True # Send immediately
		)

		frappe.log_error(f"Assignment notification sent for Incoming Document: {self.name} to {', '.join(assigned_users)}", "INCOMING DOCUMENT ASSIGNMENT NOTIFICATION SENT")

# Whitelisted functions moved to utils/sharepoint_integration.py

def create_document_version_and_notify(self, action_taken):
	"""
	Creates a new Document Version entry, fetches Sharepoint version,
	and sends email notifications.
	"""
	if not self.teams_link:
		frappe.log_warning(f"Cannot create document version for {self.name}: teams_link is empty.", "INCOMING DOCUMENT VERSIONING")
		return

	sharepoint_version = None
	try:
		# Fetch Sharepoint version using the teams_link
		sharepoint_version = get_sharepoint_version_from_link(self.teams_link)
		if not sharepoint_version:
			frappe.log_warning(f"Could not fetch Sharepoint version for link: {self.teams_link}", "INCOMING DOCUMENT VERSIONING")
			# Continue without Sharepoint version if fetching fails
	except Exception as e:
		frappe.log_error(f"Error fetching Sharepoint version for link {self.teams_link}: {e}", "INCOMING DOCUMENT VERSIONING")
		# Continue without Sharepoint version if fetching fails

	# Determine the next version number
	current_version_count = len(self.get("versions", []))
	next_version_number = current_version_count + 1

	# Create a new Document Version entry
	new_version = self.append("versions", {
		"version_number": next_version_number,
		"sharepoint_link": self.teams_link, # Use the current teams_link
		"action_taken": action_taken,
		"action_by": frappe.session.user,
		"action_timestamp": frappe.utils.now_datetime(),
		"sharepoint_version": sharepoint_version # Save the fetched Sharepoint version
		# file_url is not needed here as the file is on Sharepoint
	})

	# Save the document to persist the new child table row
	# Use ignore_permissions=True as this is a system-triggered update
	try:
		self.save(ignore_permissions=True)
		frappe.log_error(f"Created Document Version {new_version.name} for Incoming Document: {self.name}", "INCOMING DOCUMENT VERSION CREATED")
	except Exception as e:
		frappe.log_error(f"Failed to save Incoming Document {self.name} after creating version: {e}", "INCOMING DOCUMENT VERSION SAVE FAILED")
		# Decide if we should stop here or try sending emails anyway.
		# It's better to stop if the version wasn't saved.
		frappe.throw(f"Failed to save document after creating version: {e}")


	# Send email notifications
	try:
		self.notify_reviewers()
		frappe.log_error(f"Reviewer notification triggered for Incoming Document: {self.name}", "INCOMING DOCUMENT NOTIFICATION TRIGGERED")
	except Exception as e:
		frappe.log_error(f"Failed to trigger reviewer notification for Incoming Document: {self.name}: {e}", "INCOMING DOCUMENT NOTIFICATION FAILED")

	try:
		self.notify_assigned_users()
		frappe.log_error(f"Assigned users notification triggered for Incoming Document: {self.name}", "INCOMING DOCUMENT NOTIFICATION TRIGGERED")
	except Exception as e:
		frappe.log_error(f"Failed to trigger assigned users notification for Incoming Document: {self.name}: {e}", "INCOMING DOCUMENT NOTIFICATION FAILED")