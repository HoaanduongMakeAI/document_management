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

	def on_update(self):
		# Check status changes to trigger notifications
		if self.has_changed("status"):
			if self.status == "Under Review":
				self.notify_reviewers()
			elif self.status == "Tasks Assigned":
				self.notify_assigned_users()

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
				["User role", "role", "=", "Document Manager"]
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
<p><a href="{frappe.utils.get_url()}/app/incoming-document/{self.name}">Xem Văn bản đến trên ERPNext</a></p>
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
				if task.assigned_users:
					# The assigned_users field in Document Task is a Table MultiSelect,
					# which stores data as a JSON string of user emails.
					# We need to parse this JSON string and add the emails to the list.
					try:
						task_assigned_users = frappe.parse_json(task.assigned_users)
						assigned_users.extend(task_assigned_users)
					except Exception as e:
						frappe.log_error(f"Error parsing assigned_users for task {task.name}: {e}", "INCOMING DOCUMENT ASSIGNMENT NOTIFICATION FAILED")


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
				if task.assigned_users and any(user_email in assigned_users for user_email in frappe.parse_json(task.assigned_users)):
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
<p><a href="{frappe.utils.get_url()}/app/incoming-document/{self.name}">Xem Văn bản đến trên ERPNext</a></p>
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