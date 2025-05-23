# Copyright (c) 2025, Hoa An Duong and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from frappe.utils import strip_html, get_abbr
from frappe.desk.form.assign_to import add

# Import the upload function - Adjust path due to directory move
from document_management.document_management.utils.sharepoint_integration import upload_file_to_sharepoint, get_sharepoint_version_from_link

class OutgoingDocument(Document):
	# autoname is now handled by Naming Series in JSON
	# before_save hook is removed. Upload logic is now triggered manually via upload_file_via_modal.

	def before_save(self):
		# Store original approval_status before save
		if self.name: # Check if it's an existing document
			self._original_approval_status = frappe.db.get_value("Outgoing Document", self.name, "approval_status")
			self._original_teams_link = frappe.db.get_value("Outgoing Document", self.name, "teams_link") # Store original teams_link
			# Fetch and store original document tasks as a dictionary
			original_tasks = frappe.get_all("Document Task", filters={"parent": self.name}, fields=["*"])
			self._original_tasks_dict = {d.name: d for d in original_tasks} # Create dictionary for easy lookup
			# Fetch and store original document versions
			original_versions = frappe.get_all("Document Version", filters={"parent": self.name}, fields=["*"])
			self._original_versions_dict = {v.name: v for v in original_versions} # Create dictionary for easy lookup
			# frappe.msgprint(f"Original tasks dict (before_save): {self._original_tasks_dict}")
			# frappe.msgprint(f"Original versions dict (before_save): {self._original_versions_dict}")
		else: # New document
			self._original_approval_status = None
			self._original_tasks_dict = {}
			self._original_teams_link = None # Initialize for new document
			self._original_versions_dict = {} # Initialize for new document

	def on_update(self):
		# Check for changes in teams_link and create a new version if it changes
		if self.teams_link != (self._original_teams_link if hasattr(self, '_original_teams_link') else None):
			if self.teams_link: # Only trigger if teams_link is not empty after change
				# Create a new document version entry when teams_link changes
				self.create_document_version_entry("Teams Link Updated")

		# Check for new document versions and send notification
		original_versions_dict = self._original_versions_dict if hasattr(self, '_original_versions_dict') and self._original_versions_dict else {}
		current_versions_dict = {v.name: v.as_dict() for v in self.versions}

		for version_name, current_version in current_versions_dict.items():
			if version_name not in original_versions_dict:
				# This is a new version, send notification
				self.notify_version_added(frappe._dict(current_version))

		# Check for changes in document tasks and group by assignee
		original_tasks_dict = self._original_tasks_dict if hasattr(self, '_original_tasks_dict') and self._original_tasks_dict else {}
		current_tasks_dict = {d.name: d.as_dict() for d in self.document_tasks}

		# frappe.msgprint(f"Original tasks dict (on_update): {original_tasks_dict}")
		# frappe.msgprint(f"Current tasks dict (on_update): {current_tasks_dict}")

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
		# print(tasks_to_notify)

		# Add assignees from document tasks to the document's "Assign To"
		current_assignees = [d.owner for d in frappe.get_all("ToDo", filters={"reference_name": self.name, "reference_type": self.doctype}, fields=["owner"])]
		for task in self.document_tasks:
			if task.assignee and task.assignee not in current_assignees:
				try:
					args = {
						'assign_to': [task.assignee],
						'doctype': self.doctype,
						'name': self.name,
						'description': f"Công việc liên quan đến Văn bản đi: {self.outgoing_number} - {self.subject}",
					}
					add(args, ignore_permissions=True)
					frappe.log_error(f"Added assignee {task.assignee} to Outgoing Document: {self.name}", "OUTGOING DOCUMENT ASSIGNEE ADDED")
				except Exception as e:
					frappe.log_error(f"Failed to add assignee {task.assignee} to Outgoing Document: {self.name}: {e}", "OUTGOING DOCUMENT ASSIGNEE ADD FAILED")
					# Decide if we should throw an exception or just log the error.
					# For now, just log the error and continue with other tasks.
					# frappe.throw(f"Failed to add assignee {task.assignee} to Outgoing Document: {self.name}: {e}")


		# Send consolidated email to each assignee
		for assignee, tasks in tasks_to_notify.items():
			self.notify_assignee_tasks_change(assignee, tasks)

		# Handle approval workflow status changes
		self.update_approval_workflow()

	def notify_assignee_tasks_change(self, assignee, tasks):
		if not tasks:
			return

		subject = f"Cập nhật công việc liên quan đến Văn bản đi: {self.outgoing_number} - {self.subject}"
		body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Có cập nhật công việc liên quan đến văn bản đi "{self.outgoing_number} - {self.subject}":</p>
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
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
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
			frappe.log_error(f"Consolidated task change notification sent for Outgoing Document: {self.name} to {assignee}", "OUTGOING DOCUMENT CONSOLIDATED TASK NOTIFICATION SENT")
		except Exception as e:
			frappe.log_error(f"Failed to send consolidated task change notification for Outgoing Document: {self.name} to {assignee}: {e}", f"OUTGOING DOCUMENT CONSOLIDATED TASK NOTIFICATION FAILED")


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
			frappe.log_error(f"No users assigned to tasks for Outgoing Document: {self.name}", "OUTGOING DOCUMENT ASSIGNMENT NOTIFICATION FAILED")
			return

		# Construct email subject and body
		subject = f"Bạn có công việc mới liên quan đến Văn bản đi: {self.outgoing_number} - {self.subject}"
		body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Bạn được giao xử lý văn bản đi sau:</p>
<ul>
	<li><strong>Số đi:</strong> {self.outgoing_number}</li>
	<li><strong>Trích yếu:</strong> {self.subject}</li>
	<li><strong>Nơi nhận:</strong> {self.recipient}</li>
	<li><strong>Ngày ban hành:</strong> {self.date_issued}</li>
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
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
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

		frappe.log_error(f"Assignment notification sent for Outgoing Document: {self.name} to {', '.join(assigned_users)}", "OUTGOING DOCUMENT ASSIGNMENT NOTIFICATION SENT")



	def notify_version_added(self, version_data):
		"""
		Sends email notifications to assigned users when a new document version is added.
		"""
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
			frappe.log_error(f"No users assigned to tasks for Outgoing Document: {self.name} to notify about new version.", "OUTGOING DOCUMENT VERSION NOTIFICATION FAILED")
			return

		# Construct email subject and body
		subject = f"Cập nhật phiên bản mới cho Văn bản đi: {self.outgoing_number} - {self.subject}"
		body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Văn bản đi "{self.outgoing_number} - {self.subject}" đã có phiên bản mới:</p>
<ul>
	<li><strong>Phiên bản:</strong> {version_data.version_number}</li>
	<li><strong>Hành động:</strong> {version_data.action_taken}</li>
	<li><strong>Thời gian:</strong> {version_data.action_timestamp}</li>
	<li><strong>Người thực hiện:</strong> {version_data.action_by}</li>
</ul>
"""

		if version_data.sharepoint_link:
			body += f"""
<p>File phiên bản mới có thể xem tại đây:</p>
<p><a href="{version_data.sharepoint_link}">Xem File trên Microsoft Teams</a></p>
"""

		body += f"""
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết văn bản và lịch sử phiên bản:</p>
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
"""

		body += """
<p>Trân trọng,</p>
<p>Hệ thống ERPNext</p>
"""

		# Send email
		try:
			frappe.sendmail(
				recipients=assigned_users,
				subject=subject,
				message=body,
				now=True # Send immediately
			)
			frappe.log_error(f"New version notification sent for Outgoing Document: {self.name} to {', '.join(assigned_users)}", "OUTGOING DOCUMENT VERSION NOTIFICATION SENT")
		except Exception as e:
			frappe.log_error(f"Failed to send new version notification for Outgoing Document: {self.name} to {', '.join(assigned_users)}: {e}", "OUTGOING DOCUMENT VERSION NOTIFICATION FAILED")


	@frappe.whitelist()
	def create_document_version_and_notify(self, action_taken):
		"""
		Creates a new Document Version entry, fetches Sharepoint version,
		and sends email notifications.
		"""
		if not self.teams_link:
			frappe.log_warning(f"Cannot create document version for {self.name}: teams_link is empty.", "OUTGOING DOCUMENT VERSIONING")
			return

		sharepoint_version = None
		try:
			# Fetch Sharepoint version using the teams_link
			sharepoint_version = get_sharepoint_version_from_link(self.teams_link)
			if not sharepoint_version:
				frappe.log_warning(f"Could not fetch Sharepoint version for link: {self.teams_link}", "OUTGOING DOCUMENT VERSIONING")
				# Continue without Sharepoint version if fetching fails
		except Exception as e:
			frappe.log_error(f"Error fetching Sharepoint version for link {self.teams_link}: {e}", "OUTGOING DOCUMENT VERSIONING")
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
			frappe.log_error(f"Created Document Version {new_version.name} for Outgoing Document: {self.name}", "OUTGOING DOCUMENT VERSION CREATED")
		except Exception as e:
			frappe.log_error(f"Failed to save Outgoing Document {self.name} after creating version: {e}", "OUTGOING DOCUMENT VERSION SAVE FAILED")
			# It's better to stop if the version wasn't saved.
			frappe.throw(f"Failed to save document after creating version: {e}")


		# Send email notifications
		try:
			self.notify_assigned_users() # Keep this to notify assigned tasks
			frappe.log_error(f"Assigned users notification triggered for Outgoing Document: {self.name}", "OUTGOING DOCUMENT NOTIFICATION TRIGGERED")
		except Exception as e:
			frappe.log_error(f"Failed to trigger assigned users notification for Outgoing Document: {self.name}: {e}", "OUTGOING DOCUMENT NOTIFICATION FAILED")

	@frappe.whitelist()
	def create_document_version_entry(self, action_taken):
		"""
		Creates a new Document Version entry.
		This method is called from client-side or other server-side logic
		when a new version needs to be recorded (e.g., file upload, teams link change).
		"""
		if not self.teams_link:
			frappe.log_warning(f"Cannot create document version entry for {self.name}: teams_link is empty.", "OUTGOING DOCUMENT VERSIONING")
			return

		sharepoint_version = None
		try:
			# Fetch Sharepoint version using the teams_link
			sharepoint_version = get_sharepoint_version_from_link(self.teams_link)
			if not sharepoint_version:
				frappe.log_warning(f"Could not fetch Sharepoint version for link: {self.teams_link}", "OUTGOING DOCUMENT VERSIONING")
				# Continue without Sharepoint version if fetching fails
		except Exception as e:
			frappe.log_error(f"Error fetching Sharepoint version for link {self.teams_link}: {e}", "OUTGOING DOCUMENT VERSIONING")
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
			frappe.log_error(f"Created Document Version {new_version.name} for Outgoing Document: {self.name}", "OUTGOING DOCUMENT VERSION CREATED")
		except Exception as e:
			frappe.log_error(f"Failed to save Outgoing Document {self.name} after creating version: {e}", "OUTGOING DOCUMENT VERSION SAVE FAILED")
			# It's better to stop if the version wasn't saved.
			frappe.throw(f"Failed to save document after creating version: {e}")

	def update_approval_workflow(self):
		"""
		Handles state transitions and updates fields based on approval_status changes.
		"""
		original_status = self._original_approval_status if hasattr(self, '_original_approval_status') else None
		current_status = self.approval_status

		if current_status != original_status:
			frappe.log_error(f"Approval status changed from {original_status} to {current_status} for {self.name}", "OUTGOING DOCUMENT STATUS CHANGE")

			if current_status == "Pending Department Approval" and original_status == "Draft":
				# Transition from Draft to Pending Department Approval
				self.notify_department_approver_for_approval()

			elif current_status == "Department Approved" and original_status == "Pending Department Approval":
				# Transition from Pending Department Approval to Department Approved
				self.department_approver = frappe.session.user
				self.department_approval_date = frappe.utils.nowdate()
				self.notify_leadership_reviewer_for_review()

			elif current_status == "Department Rejected" and original_status == "Pending Department Approval":
				# Transition from Pending Department Approval to Department Rejected
				self.department_approver = frappe.session.user
				self.department_approval_date = frappe.utils.nowdate()
				self.notify_drafter_for_revision("Department Rejected")

			elif current_status == "Draft" and original_status in ["Department Rejected", "Leadership Rejected", "Leadership Review Request Edit"]:
				# Transition back to Draft for revision
				self.notify_drafter_for_revision(original_status)

			elif current_status == "Pending Leadership Review" and original_status == "Department Approved":
				# Transition from Department Approved to Pending Leadership Review
				self.notify_leadership_reviewer_for_review()

			elif current_status == "Leadership Approved" and original_status == "Pending Leadership Review":
				# Transition from Pending Leadership Review to Leadership Approved
				self.leadership_reviewer = frappe.session.user
				self.leadership_review_date = frappe.utils.nowdate()
				self.notify_leader_signer_for_signing()

			elif current_status == "Leadership Rejected" and original_status == "Pending Leadership Review":
				# Transition from Pending Leadership Review to Leadership Rejected
				self.leadership_reviewer = frappe.session.user
				self.leadership_review_date = frappe.utils.nowdate()
				self.notify_drafter_for_revision("Leadership Rejected")

			elif current_status == "Pending Signing" and original_status == "Leadership Approved":
				# Transition from Leadership Approved to Pending Signing
				self.notify_leader_signer_for_signing()

			elif current_status == "Signed" and original_status == "Pending Signing":
				# Transition from Pending Signing to Signed
				self.leader_signer = frappe.session.user
				self.signing_date = frappe.utils.nowdate()
				self.notify_all_parties_signed()

			elif current_status == "Issued" and original_status == "Signed":
				# Transition from Signed to Issued
				self.notify_all_parties_issued()

			elif current_status == "Rejected" and original_status not in ["Department Rejected", "Leadership Rejected"]:
				# Handle rejection from other stages if necessary
				self.notify_all_parties_rejected()

	def notify_department_approver_for_approval(self):
		"""
		Sends email notification to the department approver.
		"""
		if self.department_approver:
			subject = f"Văn bản đi chờ phê duyệt: {self.outgoing_number} - {self.subject}"
			body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Có văn bản đi chờ Anh/Chị phê duyệt:</p>
<ul>
	<li><strong>Số đi:</strong> {self.outgoing_number}</li>
	<li><strong>Trích yếu:</strong> {self.subject}</li>
	<li><strong>Bộ phận soạn thảo:</strong> {self.created_by_user}</li>
</ul>
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết và thực hiện phê duyệt:</p>
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
"""
			try:
				frappe.sendmail(
					recipients=[self.department_approver],
					subject=subject,
					message=body,
					now=True
				)
				frappe.log_error(f"Notification sent to department approver {self.department_approver} for {self.name}", "OUTGOING DOCUMENT NOTIFICATION")
			except Exception as e:
				frappe.log_error(f"Failed to send notification to department approver {self.department_approver} for {self.name}: {e}", "OUTGOING DOCUMENT NOTIFICATION FAILED")

	def notify_leadership_reviewer_for_review(self):
		"""
		Sends email notification to the leadership reviewer.
		"""
		if self.leadership_reviewer:
			subject = f"Văn bản đi chờ xem xét: {self.outgoing_number} - {self.subject}"
			body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Có văn bản đi chờ Anh/Chị xem xét:</p>
<ul>
	<li><strong>Số đi:</strong> {self.outgoing_number}</li>
	<li><strong>Trích yếu:</strong> {self.subject}</li>
	<li><strong>Bộ phận soạn thảo:</strong> {self.created_by_user}</li>
	<li><strong>Bộ phận phê duyệt:</strong> {self.department}</li>
</ul>
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết và thực hiện xem xét:</p>
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
"""
			try:
				frappe.sendmail(
					recipients=[self.leadership_reviewer],
					subject=subject,
					message=body,
					now=True
				)
				frappe.log_error(f"Notification sent to leadership reviewer {self.leadership_reviewer} for {self.name}", "OUTGOING DOCUMENT NOTIFICATION")
			except Exception as e:
				frappe.log_error(f"Failed to send notification to leadership reviewer {self.leadership_reviewer} for {self.name}: {e}", "OUTGOING DOCUMENT NOTIFICATION FAILED")

	def notify_leader_signer_for_signing(self):
		"""
		Sends email notification to the leader signer.
		"""
		if self.leader_signer:
			subject = f"Văn bản đi chờ ký duyệt: {self.outgoing_number} - {self.subject}"
			body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Có văn bản đi chờ Anh/Chị ký duyệt:</p>
<ul>
	<li><strong>Số đi:</strong> {self.outgoing_number}</li>
	<li><strong>Trích yếu:</strong> {self.subject}</li>
	<li><strong>Bộ phận soạn thảo:</strong> {self.created_by_user}</li>
	<li><strong>Bộ phận phê duyệt:</strong> {self.department}</li>
	<li><strong>Người xem xét:</strong> {self.leadership_reviewer}</li>
</ul>
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết và thực hiện ký duyệt:</p>
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
"""
			try:
				frappe.sendmail(
					recipients=[self.leader_signer],
					subject=subject,
					message=body,
					now=True
				)
				frappe.log_error(f"Notification sent to leader signer {self.leader_signer} for {self.name}", "OUTGOING DOCUMENT NOTIFICATION")
			except Exception as e:
				frappe.log_error(f"Failed to send notification to leader signer {self.leader_signer} for {self.name}: {e}", "OUTGOING DOCUMENT NOTIFICATION FAILED")

	def notify_drafter_for_revision(self, rejection_reason):
		"""
		Sends email notification to the drafter for revision.
		"""
		if self.created_by_user:
			subject = f"Văn bản đi cần chỉnh sửa: {self.outgoing_number} - {self.subject}"
			body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Văn bản đi "{self.outgoing_number} - {self.subject}" cần được chỉnh sửa.</p>
<p><strong>Lý do:</strong> {rejection_reason}</p>
"""
			if rejection_reason == "Department Rejected" and self.department_approval_feedback:
				body += f"<p><strong>Ý kiến phê duyệt Bộ phận:</strong> {self.department_approval_feedback}</p>"
			elif rejection_reason in ["Leadership Rejected", "Leadership Review Request Edit"] and self.leadership_review_feedback:
				body += f"<p><strong>Ý kiến xem xét Lãnh đạo:</strong> {self.leadership_review_feedback}</p>"

			body += f"""
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết và thực hiện chỉnh sửa:</p>
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
"""
			try:
				frappe.sendmail(
					recipients=[self.created_by_user],
					subject=subject,
					message=body,
					now=True
				)
				frappe.log_error(f"Notification sent to drafter {self.created_by_user} for revision of {self.name}", "OUTGOING DOCUMENT NOTIFICATION")
			except Exception as e:
				frappe.log_error(f"Failed to send notification to drafter {self.created_by_user} for revision of {self.name}: {e}", "OUTGOING DOCUMENT NOTIFICATION FAILED")


	def notify_all_parties_signed(self):
		"""
		Sends email notification to all relevant parties after signing.
		"""
		recipients = [self.created_by_user, self.department_approver, self.leadership_reviewer]
		recipients = list(set([r for r in recipients if r])) # Remove duplicates and None

		if recipients:
			subject = f"Văn bản đi đã được ký duyệt: {self.outgoing_number} - {self.subject}"
			body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Văn bản đi "{self.outgoing_number} - {self.subject}" đã được ký duyệt.</p>
<ul>
	<li><strong>Người ký duyệt:</strong> {self.leader_signer}</li>
	<li><strong>Ngày ký duyệt:</strong> {self.signing_date}</li>
</ul>
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết:</p>
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
"""
			try:
				frappe.sendmail(
					recipients=recipients,
					subject=subject,
					message=body,
					now=True
				)
				frappe.log_error(f"Notification sent to all parties for signed document {self.name}", "OUTGOING DOCUMENT NOTIFICATION")
			except Exception as e:
				frappe.log_error(f"Failed to send notification to all parties for signed document {self.name}: {e}", "OUTGOING DOCUMENT NOTIFICATION FAILED")


	def notify_all_parties_issued(self):
		"""
		Sends email notification to all relevant parties after issuing.
		"""
		recipients = [self.created_by_user, self.department_approver, self.leadership_reviewer, self.leader_signer]
		recipients = list(set([r for r in recipients if r])) # Remove duplicates and None

		if recipients:
			subject = f"Văn bản đi đã được ban hành: {self.outgoing_number} - {self.subject}"
			body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Văn bản đi "{self.outgoing_number} - {self.subject}" đã được ban hành.</p>
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết:</p>
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
"""
			try:
				frappe.sendmail(
					recipients=recipients,
					subject=subject,
					message=body,
					now=True
				)
				frappe.log_error(f"Notification sent to all parties for issued document {self.name}", "OUTGOING DOCUMENT NOTIFICATION")
			except Exception as e:
				frappe.log_error(f"Failed to send notification to all parties for issued document {self.name}: {e}", "OUTGOING DOCUMENT NOTIFICATION FAILED")


	def notify_all_parties_rejected(self):
		"""
		Sends email notification to all relevant parties when rejected.
		"""
		recipients = [self.created_by_user, self.department_approver, self.leadership_reviewer, self.leader_signer]
		recipients = list(set([r for r in recipients if r])) # Remove duplicates and None

		if recipients:
			subject = f"Văn bản đi đã bị từ chối: {self.outgoing_number} - {self.subject}"
			body = f"""
<p>Kính gửi Anh/Chị,</p>
<p>Văn bản đi "{self.outgoing_number} - {self.subject}" đã bị từ chối.</p>
"""
			# Include feedback if available
			if self.department_approval_feedback:
				body += f"<p><strong>Ý kiến phê duyệt Bộ phận:</strong> {self.department_approval_feedback}</p>"
			if self.leadership_review_feedback:
				body += f"<p><strong>Ý kiến xem xét Lãnh đạo:</strong> {self.leadership_review_feedback}</p>"

			body += f"""
<p>Vui lòng truy cập vào hệ thống ERPNext để xem chi tiết:</p>
<p><a href="/app/outgoing-document/{self.name}">Xem Văn bản đi trên ERPNext</a></p>
"""
			try:
				frappe.sendmail(
					recipients=recipients,
					subject=subject,
					message=body,
					now=True
				)
				frappe.log_error(f"Notification sent to all parties for rejected document {self.name}", "OUTGOING DOCUMENT NOTIFICATION")
			except Exception as e:
				frappe.log_error(f"Failed to send notification to all parties for rejected document {self.name}: {e}", "OUTGOING DOCUMENT NOTIFICATION FAILED")

	@frappe.whitelist()
	def submit_for_department_approval(self):
		"""
		Sets the status to Pending Department Approval and notifies the department approver.
		Only allowed if status is Draft and current user is the creator.
		"""
		if self.approval_status != "Draft" or self.created_by_user != frappe.session.user:
			frappe.throw("You are not allowed to submit this document for department approval.")

		self.approval_status = "Pending Department Approval"
		self.save()
		self.notify_department_approver_for_approval()

	@frappe.whitelist()
	def department_approve(self):
		"""
		Sets the status to Department Approved, records approver and date, and notifies leadership.
		Only allowed if status is Pending Department Approval and current user is the department approver.
		"""
		if self.approval_status != "Pending Department Approval" or self.department_approver != frappe.session.user:
			frappe.throw("You are not allowed to approve this document at the department level.")

		self.approval_status = "Department Approved"
		self.department_approver = frappe.session.user
		self.department_approval_date = frappe.utils.nowdate()
		self.save()
		self.notify_leadership_reviewer_for_review()

	@frappe.whitelist()
	def department_reject(self):
		"""
		Sets the status to Department Rejected, records approver and date, and notifies the drafter.
		Only allowed if status is Pending Department Approval and current user is the department approver.
		"""
		if self.approval_status != "Pending Department Approval" or self.department_approver != frappe.session.user:
			frappe.throw("You are not allowed to reject this document at the department level.")

		self.approval_status = "Department Rejected"
		self.department_approver = frappe.session.user
		self.department_approval_date = frappe.utils.nowdate()
		self.save()
		self.notify_drafter_for_revision("Department Rejected")

	@frappe.whitelist()
	def department_request_edit(self):
		"""
		Sets the status back to Draft, records approver and date, and notifies the drafter for revision.
		Only allowed if status is Pending Department Approval and current user is the department approver.
		"""
		if self.approval_status != "Pending Department Approval" or self.department_approver != frappe.session.user:
			frappe.throw("You are not allowed to request edits for this document at the department level.")

		self.approval_status = "Draft"
		self.department_approver = frappe.session.user
		self.department_approval_date = frappe.utils.nowdate()
		self.save()
		self.notify_drafter_for_revision("Department Review Request Edit")

	@frappe.whitelist()
	def leadership_approve(self):
		"""
		Sets the status to Leadership Approved, records reviewer and date, and notifies the signer.
		Only allowed if status is Pending Leadership Review and current user is the leadership reviewer.
		"""
		if self.approval_status != "Pending Leadership Review" or self.leadership_reviewer != frappe.session.user:
			frappe.throw("You are not allowed to approve this document at the leadership level.")

		self.approval_status = "Leadership Approved"
		self.leadership_reviewer = frappe.session.user
		self.leadership_review_date = frappe.utils.nowdate()
		self.save()
		self.notify_leader_signer_for_signing()

	@frappe.whitelist()
	def leadership_reject(self):
		"""
		Sets the status to Leadership Rejected, records reviewer and date, and notifies the drafter.
		Only allowed if status is Pending Leadership Review and current user is the leadership reviewer.
		"""
		if self.approval_status != "Pending Leadership Review" or self.leadership_reviewer != frappe.session.user:
			frappe.throw("You are not allowed to reject this document at the leadership level.")

		self.approval_status = "Leadership Rejected"
		self.leadership_reviewer = frappe.session.user
		self.leadership_review_date = frappe.utils.nowdate()
		self.save()
		self.notify_drafter_for_revision("Leadership Rejected")

	@frappe.whitelist()
	def leadership_request_edit(self):
		"""
		Sets the status back to Draft, records reviewer and date, and notifies the drafter for revision.
		Only allowed if status is Pending Leadership Review and current user is the leadership reviewer.
		"""
		if self.approval_status != "Pending Leadership Review" or self.leadership_reviewer != frappe.session.user:
			frappe.throw("You are not allowed to request edits for this document at the leadership level.")

		self.approval_status = "Draft"
		self.leadership_reviewer = frappe.session.user
		self.leadership_review_date = frappe.utils.nowdate()
		self.save()
		self.notify_drafter_for_revision("Leadership Review Request Edit")

	@frappe.whitelist()
	def sign_document(self):
		"""
		Sets the status to Signed, records signer and date, and notifies all parties.
		Only allowed if status is Pending Signing and current user is the leader signer.
		"""
		if self.approval_status != "Pending Signing" or self.leader_signer != frappe.session.user:
			frappe.throw("You are not allowed to sign this document.")

		self.approval_status = "Signed"
		self.leader_signer = frappe.session.user
		self.signing_date = frappe.utils.nowdate()
		self.save()
		self.notify_all_parties_signed()