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
		# Check if the status has changed to "Pending Review"
		# This hook is triggered on save, check for workflow transition to Pending Review
		if self.has_workflow_permission('Submit for Review') and self.docstatus == 0 and self.status == 'Pending Review':
			self.notify_reviewers()

	def on_transition(self, workflow_action):
		# This method is called by the workflow engine on each transition
		if workflow_action == 'Assign for Processing':
			self.notify_assigned_users()


	def notify_reviewers(self):
		# Get users with 'Lãnh đạo' or 'Cố vấn' roles
		# NOTE: Role names 'Lãnh đạo' and 'Cố vấn' are based on the analysis document.
		# Please adjust if the actual role names in the system are different.
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
		# Get users assigned to this document via the ToDo (Assignment) doctype
		assigned_users = frappe.get_all("ToDo",
			filters={
				"reference_doctype": self.doctype,
				"reference_name": self.name,
				"status": "Open" # Only notify for open assignments
			},
			pluck="owner" # The 'owner' field in ToDo is the assigned user
		)
		

		if not assigned_users:
			frappe.log_error(f"No users assigned to Incoming Document: {self.name}", "INCOMING DOCUMENT ASSIGNMENT NOTIFICATION FAILED")
			return

		# Construct email subject and body
		subject = f"Bạn có văn bản đến cần xử lý: {self.incoming_number} - {self.subject}"
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
<p><strong>Ý kiến chỉ đạo/Hướng dẫn xử lý:</strong></p>
<p>{self.instructions}</p>
"""

		body += f"""
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết và cập nhật tiến độ:</p>
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