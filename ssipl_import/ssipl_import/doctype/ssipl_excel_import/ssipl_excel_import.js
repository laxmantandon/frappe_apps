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
	},
	start_import_2: function(frm) {
		if(!frm.doc.excel_file_2) {
			frappe.throw(__("Please select excel sheet"));
		}

		frm.call("start_import", {
			excel_file: frm.doc.excel_file_2
		}, () => {
			// call back do somthing here
		});
	},
	start_import_3: function(frm) {
		if(!frm.doc.excel_file_3) {
			frappe.throw(__("Please select excel sheet"));
		}

		frm.call("start_import", {
			excel_file: frm.doc.excel_file_3
		}, () => {
			// call back do somthing here
		});
	},
	start_import_4: function(frm) {
		if(!frm.doc.excel_file_4) {
			frappe.throw(__("Please select excel sheet"));
		}

		frm.call("start_import", {
			excel_file: frm.doc.excel_file_4
		}, () => {
			// call back do somthing here
		});
	},
	start_import_5: function(frm) {
		if(!frm.doc.excel_file_5) {
			frappe.throw(__("Please select excel sheet"));
		}

		frm.call("start_import", {
			excel_file: frm.doc.excel_file_5
		}, () => {
			// call back do somthing here
		});
	}
});
