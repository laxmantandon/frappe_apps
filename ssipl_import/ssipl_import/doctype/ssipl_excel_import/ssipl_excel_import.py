# -*- coding: utf-8 -*-
# Copyright (c) 2020, Laxman and contributors
# For license information, please see license.txt

from __future__ import unicode_literals
import frappe
import xlrd
import datetime
from frappe.model.document import Document

class SSIPLExcelImport(Document):
	
	def start_import(self):

		excel_file = frappe.local.site_path + self.excel_file
		wb = xlrd.open_workbook(excel_file)
		sheet = wb.sheet_by_index(0)

		self.import_items()

		
	def import_items(self):

		excel_file = frappe.local.site_path + self.excel_file
		wb = xlrd.open_workbook(excel_file)
		sheet = wb.sheet_by_index(0)
		
		for i in range(sheet.nrows):
			if i > 0:
				self.create_item_groups(i, sheet)
				self.create_item_subgroups(i, sheet)
				self.create_suppliers(i, sheet)
				self.create_brands(i, sheet)
				self.create_items(i, sheet)
				self.create_price_lists(i, sheet)
		
		self.create_stock_entry(sheet)
		frappe.msgprint("File Processed Successfully")

	def create_item_groups(self, i, sheet):

		item_group_name = sheet.cell_value(i, 13)
		if not frappe.db.exists('Item Group', item_group_name):

			item_group = frappe.get_doc({
					"name": item_group_name,
					"parent": "All Item Groups",
					"item_group_name": item_group_name,
					"parent_item_group": "All Item Groups",
					"is_group": 1,
					"doctype": "Item Group",
					})
			item_group.insert()

	def create_item_subgroups(self, i, sheet):

		item_sub_group_name = sheet.cell_value(i, 14) # + ' - ' + sheet.cell_value(i, 13)
		item_parent_group_name = sheet.cell_value(i, 13)
		if not frappe.db.exists('Item Group', item_sub_group_name):

			item_group = frappe.get_doc({
					"name": item_sub_group_name,
					"parent": "All Item Groups",
					"item_group_name": item_sub_group_name,
					"parent_item_group": item_parent_group_name,
					"is_group": 0,
					"doctype": "Item Group",
					})
			item_group.insert()

	def create_suppliers(self, i, sheet):
		
		supplier_name = sheet.cell_value(i, 4)
		if not frappe.db.exists('Supplier', supplier_name):

			supplier = frappe.get_doc({
					"name": supplier_name,
					"supplier_name": supplier_name,
					"country": "India",
					"supplier_group": "All Supplier Groups",
					"supplier_type": "Company",
					"language": "en",
					"doctype": "Supplier"
					})
			supplier.insert()

	def create_brands(self, i, sheet):

		brand_name = sheet.cell_value(i, 23)
		if not frappe.db.exists('Brand', brand_name):

			brand = frappe.get_doc({
					"name": brand_name,
					"brand": brand_name,
					"doctype": "Brand"
					})
			brand.insert()		

	def create_items(self, i, sheet):

		item_name = sheet.cell_value(i, 2)
		item_desc = sheet.cell_value(i, 3)
		item_group = sheet.cell_value(i, 14)
		gst_hsn_code = sheet.cell_value(i, 33)
		item_valuation = sheet.cell_value(i, 5)
		standard_rate = sheet.cell_value(i, 8)
		brand = sheet.cell_value(i, 23)
		supplier = sheet.cell_value(i, 4)
		category_type = sheet.cell_value(i, 21)
		design_name = sheet.cell_value(i, 24)
		size_name = sheet.cell_value(i, 31)
		color_name = sheet.cell_value(i, 32)
		gst_local = self.get_local_tax_template(sheet.cell_value(i, 34))
		gst_interstate = self.get_interstate_tax_template(sheet.cell_value(i, 34))

		if not frappe.db.exists('Item', item_name):

			item = frappe.get_doc(
				{
					"name": item_name,
					"item_code": item_name,
					"item_name": item_desc,
					"item_group": item_group,
					"category_type": category_type,
					"design_name": design_name,
					"size_name": size_name,
					"color_name": color_name,
					"gst_hsn_code": gst_hsn_code,
					"stock_uom":"Nos",
					"is_stock_item":1,
					"include_item_in_manufacturing":1,
					"valuation_rate": item_valuation,
					"standard_rate": standard_rate,
					"brand": brand,
					"description": item_desc,
					"default_material_request_type":"Purchase",
					"valuation_method":"FIFO",
					"has_batch_no":1,
					"create_new_batch":1,
					"batch_number_series":"SSIPL.#####",
					"variant_based_on":"Item Attribute",
					"is_purchase_item":1,
					"country_of_origin":"India",
					"is_sales_item":1,
					"doctype":"Item",
					"barcodes":[
						{
							"barcode": item_name,
							"barcode_type":"",
							"doctype":"Item Barcode",
							"name": item_name,
							"parent": item_name,
							"parentfield":"barcodes",
							"parenttype":"Item"
						}
					],
					"uoms":[
						{
							"parent": item_name,
							"parentfield":"uoms",
							"parenttype":"Item",
							"uom":"Nos",
							"conversion_factor":1,
							"doctype":"UOM Conversion Detail"
						}
					],
					"supplier_items":[
						{
							"parent":item_name,
							"parentfield":"supplier_items",
							"parenttype":"Item",
							"supplier": supplier,
							"doctype":"Item Supplier"
						}
					],
					"taxes":[
						{
							"parent": item_name,
							"parentfield":"taxes",
							"parenttype":"Item",
							"item_tax_template": gst_interstate,
							"tax_category":"INTERSTATE",
							"valid_from":"2020-04-01",
							"doctype":"Item Tax"
						},
						{
							"parent": item_name,
							"parentfield":"taxes",
							"parenttype":"Item",
							"item_tax_template": gst_local,
							"tax_category":"LOCAL",
							"valid_from":"2020-04-01",
							"doctype":"Item Tax"
						}
					]
					})
			item.insert()		
		
	def get_local_tax_template(self, taxper):
		# frappe.msgpriknt(taxper, 'Local')
		tax_slab = {
			"0": "GST TAX FREE - LOCAL",
			"5": "GST @ 5% - LOCAL",
			"12": "GST @ 12% - LOCAL",
			"18": "GST @ 18% - LOCAL",
			"28": "GST @ 28% - LOCAL"
		}

		return tax_slab.get(taxper, "")

	def get_interstate_tax_template(self, taxper):
		# frappe.msgprint(taxper, 'Interstate')
		tax_slab = {
			"0": "GST TAX FREE - INTERSTATE",
			"5": "GST @ 5% - INTERSTATE",
			"12": "GST @ 12% - INTERSTATE",
			"18": "GST @ 18% - INTERSTATE",
			"28": "GST @ 28% - INTERSTATE"
		}

		return tax_slab.get(taxper, "")

	def create_price_lists(self, i, sheet):

		item_name = sheet.cell_value(i, 2)
		item_desc = sheet.cell_value(i, 3)
		brand = sheet.cell_value(i, 23)
		wholesale_price = sheet.cell_value(i, 26)
		retail_price = sheet.cell_value(i, 8)
		dummy_price = sheet.cell_value(i, 20)
		franchise_price = sheet.cell_value(i, 47)
		valid_from = datetime.datetime.strptime(sheet.cell_value(i, 1), '%d-%m-%Y')

		# if not frappe.db.count('Item Price', {'item_code': item_code, 'price_list': 'Wholesale', 'valid_from': valid_from}) > 0:

		if not frappe.db.count('Item Price', {'item_code': item_name, 'price_list': 'Wholesale', 'valid_from': valid_from.date()}) > 0:
		
			wholesale_item_price = frappe.get_doc({
					"item_code": item_name,
					"item_name": item_desc,
					"brand": brand,
					"item_description": item_desc,
					"price_list": "Wholesale",
					"buying": 0,
					"selling": 1,
					"currency": "INR",
					"price_list_rate": wholesale_price,
					"valid_from": valid_from.date(),
					"lead_time_days": 0,
					"doctype": "Item Price",
					"note": "Imported from ssipl import module"
					})
			wholesale_item_price.insert()

		if not frappe.db.count('Item Price', {'item_code': item_name, 'price_list': 'Retail', 'valid_from': valid_from.date()}) > 0:
		
			retail_item_price = frappe.get_doc({
					"item_code": item_name,
					"item_name": item_desc,
					"brand": brand,
					"item_description": item_desc,
					"price_list": "Retail",
					"buying": 0,
					"selling": 1,
					"currency": "INR",
					"price_list_rate": retail_price,
					"valid_from": valid_from.date(),
					"lead_time_days": 0,
					"doctype": "Item Price",
					"note": "Imported from ssipl import module"
					})
			retail_item_price.insert()

		if not frappe.db.count('Item Price', {'item_code': item_name, 'price_list': 'Franchise', 'valid_from': valid_from.date()}) > 0:
		
			franchise_item_price = frappe.get_doc({
					"item_code": item_name,
					"item_name": item_desc,
					"brand": brand,
					"item_description": item_desc,
					"price_list": "Franchise",
					"buying": 0,
					"selling": 1,
					"currency": "INR",
					"price_list_rate": franchise_price,
					"valid_from": valid_from.date(),
					"lead_time_days": 0,
					"doctype": "Item Price",
					"note": "Imported from ssipl import module"
					})
			franchise_item_price.insert()

	def create_stock_entry(self, sheet):

		posting_date = datetime.datetime.strptime(sheet.cell_value(1, 1), '%d-%m-%Y')
		ssipl_bill_no = sheet.cell_value(1, 0)
		
		items = []
		for i in range(sheet.nrows):
			if i > 0: 
				item_code = sheet.cell_value(i, 2)
				item_group = sheet.cell_value(i, 14)
				item_name = sheet.cell_value(i, 3)
				retail_price = sheet.cell_value(i, 8)
				qty = float(sheet.cell_value(i, 7))
				basic_rate = sheet.cell_value(i, 5)
				basic_amount = float(retail_price) * float(sheet.cell_value(i, 7))

				items.append({
					"parentfield": "items",
					"parenttype": "Stock Entry",
					"t_warehouse": self.warehouse,
					"item_code": item_code,
					"item_group": item_group,
					"item_name": item_name,
					"description": item_name,
					"qty": qty,
					"basic_rate": basic_rate,
					"basic_amount": basic_amount,
					"amount": basic_amount,
					"valuation_rate": basic_rate,
					"uom": "Nos",
					"conversion_factor": 1,
					"stock_uom": "Nos",
					"transfer_qty": qty,
					# "expense_account": "Stock Adjustment - TC",
					# "cost_center": "Main - TC",
					"doctype": "Stock Entry Detail"
				})
				

		if not frappe.db.count('Stock Entry', {'ssipl_bill_no': ssipl_bill_no}) > 0:
			
			stock_entry = frappe.get_doc({
					"title": "Material Receipt",
					"naming_series": "MAT-STE-.YYYY.-",
					"stock_entry_type": "Material Receipt",
					"purpose": "Material Receipt",
					"company": self.company,
					"posting_date": posting_date.date(),
					"use_multi_level_bom": 1,
					"to_warehouse": self.warehouse,
					"ssipl_bill_no": ssipl_bill_no,
					# "total_incoming_value": ,
					# "value_difference": ,
					# "total_amount": ,
					"doctype": "Stock Entry",
					"items": items
				})
			stock_entry.insert()
			stock_entry.submit()
			self.stock_entry = stock_entry.name