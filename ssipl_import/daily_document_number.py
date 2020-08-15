# -*- coding: utf-8 -*-
# Copyright (c) 2020, Laxman and contributors
# For license information, please see license.txt

from __future__ import unicode_literals
import frappe
import datetime
from frappe.model.naming import make_autoname
from frappe.utils import date_diff

@frappe.whitelist()
def set_auto_number(doc, method):
    start_date = frappe.defaults.get_user_default("year_start_date")
    dayofyear = date_diff(doc.posting_date, start_date)
    doc.name = make_autoname(doc.naming_series + '-.' + str(dayofyear) + '.-' + '.#####')