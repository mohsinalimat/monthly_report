frappe.query_reports["Monthly Income Statement"] = {
	"filters": [
		{
			fieldname: 'company',
			label: __('Company'),
			fieldtype: 'Link',
			options: 'Company',
			default: frappe.defaults.get_user_default('company'),
			hidden: true,
		},
		{
			fieldname: "finance_book",
			label: __("Finance Book"),
			fieldtype: "Link",
			options: "Finance Book",
			hidden: true,
		},
		{
			fieldname: "filter_based_on",
			label: __("Filter Based On"),
			fieldtype: "Select",
			options: ["Fiscal Year", "Date Range"],
			default: ["Fiscal Year"],
			reqd: true,
			on_change: function () {
				let filter_based_on = frappe.query_reports.get_filter_value('filter_based_on');
				frappe.query_reports.toggle_filter_display('from_fiscal_year', filter_based_on === 'Date Range');
				frappe.query_reports.toggle_filter_display('to_fiscal_year', filter_based_on === 'Date Range');
				frappe.query_reports.toggle_filter_display('period_start_date', filter_based_on === 'Fiscal Year');
				frappe.query_reports.toggle_filter_display('period_end_date', filter_based_on === 'Fiscal Year');
				frappe.query_reports.refresh();
			},
			hidden: true,
		},
		{
			fieldname: "print_group",
			label: __("Print Group"),
			fieldtype: "Select",
			options: "30 - Trade Sales",
			default: "30 - Trade Sales",
			reqd: true,
			hidden: true,
		},
		{
			"fieldname": "period_end_month",
			"fieldtype": "Select",
			"label": "Month",
			"options": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
			"default": "January",
			"mandatory": 0,
			"wildcard_filter": 0
		},
		// {
		// 	"fieldname": "period_start_date",
		// 	"label": __("Start Date"),
		// 	"fieldtype": "Date",
		// 	"default": new Date(2021, 6, 1),
		// 	"reqd": true,
		// 	"depends_on": "eval:doc.filter_based_on == 'Date Range'",
		// },
		// {
		// 	"fieldname": "period_end_date",
		// 	"label": __("End Date"),
		// 	"fieldtype": "Date",
		// 	"default": moment(),
		// 	"reqd": true,
		// 	"depends_on": "eval:doc.filter_based_on == 'Date Range'",
		// 	hidden: true,
		// },
		{
			"fieldname": "from_fiscal_year",
			"label": __("Start Year"),
			"fieldtype": "Link",
			"options": "Fiscal Year",
			"default": frappe.defaults.get_user_default("fiscal_year"),
			"reqd": 1,
			"depends_on": "eval:doc.filter_based_on == 'Fiscal Year'"
		},
		{
			"fieldname": "to_fiscal_year",
			"label": __("End Year"),
			"fieldtype": "Link",
			"options": "Fiscal Year",
			"default": frappe.defaults.get_user_default("fiscal_year"),
			"reqd": 1,
			"depends_on": "eval:doc.filter_based_on == 'Fiscal Year'"
		},
		{
			"fieldname": "periodicity",
			"label": __("Periodicity"),
			"fieldtype": "Select",
			"options": [
				{ "value": "Monthly", "label": __("Monthly") }
			],
			"default": "Monthly",
			"reqd": true,
			"hidden": true,
		},
		{
			"fieldname": "cost_center",
			"label": __("Cost Center"),
			"fieldtype": "MultiSelectList",
			get_data: function (txt) {
				return frappe.db.get_link_options('Cost Center', txt, {
					company: frappe.query_report.get_filter_value("company")
				});
			},
		},
	]
}
