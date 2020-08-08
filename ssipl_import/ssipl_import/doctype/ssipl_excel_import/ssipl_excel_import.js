// Copyright (c) 2020, Laxman and contributors
// For license information, please see license.txt

frappe.ui.form.on('SSIPL Excel Import', {
	// refresh: function(frm) {

	// }
	start_import: function(frm) {
		if(!frm.doc.excel_file) {
			frappe.throw(__("Please select excel sheet"));
		}

		frm.call("start_import", {
			excel_file: frm.doc.excel_file
		}, () => {
			// call back do somthing here
		});
	}
});
