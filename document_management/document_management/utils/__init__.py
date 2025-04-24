import frappe

# Helper function to easily get settings
# Use frappe.cache() for better performance
@frappe.whitelist()
def get_settings():
	# Use frappe.get_cached_doc to handle caching and retrieval of single doctypes
	return frappe.get_cached_doc("Document Management Settings")