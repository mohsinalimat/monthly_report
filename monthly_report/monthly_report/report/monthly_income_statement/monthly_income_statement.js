frappe.query_reports["Monthly Income Statement"] = {
	"filters": [
		{"fieldname": 'company',          "label": "Company",         "fieldtype": 'Link',   "options": 'Company', "default": frappe.defaults.get_user_default('company'), "hidden": true,},
		{"fieldname": "finance_book",     "label": "Finance Book",    "fieldtype": "Link",   "options": "Finance Book", "hidden": true,},
		{"fieldname": "to_fiscal_year",   "label": "End Year",        "fieldtype": "Link",   "options": "Fiscal Year", "default": frappe.defaults.get_user_default("fiscal_year")-1, "reqd": 1, "depends_on": "eval:doc.filter_based_on == 'Fiscal Year'"},
		{"fieldname": "print_group",      "label": "Print Group",     "fieldtype": "Select", "options": "30 - Trade Sales", "default": "30 - Trade Sales", "reqd": true, "hidden": true},
		{"fieldname": "period_end_month", "label": "Month",           "fieldtype": "Select", "options": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], "default": "January", "mandatory": 0, "wildcard_filter": 0},
		{"fieldname": "periodicity",      "label": "Periodicity",     "fieldtype": "Select", "options": [{ "value": "Monthly", "label": __("Monthly") }], "default": "Monthly", "reqd": true, "hidden": true,},
		{"fieldname": "filter_based_on",  "label": "Filter Based On", "fieldtype": "Select", "options": ["Fiscal Year", "Date Range"], "default": ["Fiscal Year"], "reqd": true, "hidden": true},
		{"fieldname": "cost_center",      "label": "Cost Center",     "fieldtype": "MultiSelectList", get_data: function (txt) {return frappe.db.get_link_options('Cost Center', txt, {company: frappe.query_report.get_filter_value("company")});}},
	],

	onload: function(report) {
		report.page.add_inner_button(__("Export Report"), function () {
			debugger
			let filters = report.get_values();

			frappe.call({
				method: 'monthly_report.monthly_report.report.monthly_income_statement.monthly_income_statement.get_records',
				args: {
					report_name: report.report_name,
					filters: filters
				},

				callback: function (r) {
					$(".report-wrapper").html("")
					$(".justify-center").remove()

					if (r.message[1] != "") {
						generate_table(r.message, filters.company, filters.period_end_month, filters.to_fiscal_year)
					} else {
						alert("No record found.")
					}
				}
			});
		});
	},
}

// ======================================================================
// GENERATOR FUNCTIONS
// ======================================================================

function generate_table(message, company, month, year) {
	// tbh i dont know what these are, just that things break without them 
	var $export_id  = "export_id";
	var table_count = [];
	table_count[0]  = "#" + $export_id;
	console.log(message);

	var table_html = "";
	var table_css  = "";
	var month_name = (month.slice(0, 3)).toLowerCase();
	var curr_month_year = month_name + "_" + year;
	var prev_month_year = month_name + "_" + (parseInt(year) - 1).toString();

	// css for the table 
	table_css  = generate_table_css();
	
	// the table containing all the data in html format
	table_html += '<div id="data">';
	table_html += '    <table style="font-weight: normal; font-family: Calibri; font-size: 10pt;" id=' + $export_id + '>';
	table_html += generate_table_caption(company, month, year)
	table_html += generate_table_head(month, year)
	table_html += generate_table_body(message, curr_month_year, prev_month_year)
	table_html += '    </table>';
	table_html += '</div>';

	// append the css & html, then export as xls
	$(".report-wrapper").hide();
	$(".report-wrapper").append(table_css);
	$(".report-wrapper").append(table_html);
	tables_to_excel(table_count, curr_month_year +'.xls');
}

function generate_table_body(message, curr_month_year, prev_month_year) {
	// variables that store each row's data
	var account    = ""; // name of the print group
	var is_group   = ""; // name of the print group
	var curr_data  = ""; // current fiscal year's data
	var prev_data  = ""; // previous fiscal year's total
	var curr_ytd   = ""; // current fiscal year's data
	var prev_ytd   = ""; // preious fiscal year's total
	var parent     = "";
	var last_parent = "";

	var table_body = ""; // holds the html that is returned
			
	// variables for formatting and calculations
	var indent = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";  // prints an indent
	// var category_names  = ["Other Income", "Product Sales", "Operating Expenses", "Other Expenses"];
	var total_other_income = get_category_total("Other Income", message, curr_month_year, prev_month_year);
	var total_product_sales = get_category_total("Product Sales", message, curr_month_year, prev_month_year);
	var total_operating_expenses = get_category_total("Operating Expenses", message, curr_month_year, prev_month_year);
	var total_other_expenses = get_category_total("Other Expenses", message, curr_month_year, prev_month_year);

	table_body += '<tbody>'; // start html table body
	for (let i = 0; i < message[1].length; i++) {
		
		// check for indentation and print when needed
		if (message[1][i].indent == 0) account = message[1][i]["account"];
		if (message[1][i].indent == 1) account = indent + message[1][i]["account"];
		if (message[1][i].indent == 2) account = indent + indent + message[1][i]["account"];
		
		if (message[1][i]["account"] != "") {
			curr_data = message[1][i][curr_month_year];
			prev_data = message[1][i][prev_month_year];
			curr_ytd  = message[1][i]["total"];
			prev_ytd  = message[1][i]["prev_year_total"];
			is_group  = message[1][i]["is_group"];
			parent    = message[1][i]["parent_account"];
		}


		if (is_group) { 
			// these rows should not contain any values

			// if (parent == "Income - WW" || parent == "Expenses - WW") {
			// 	if (message[1][i]["account"] == "Product Sales") {
			// 		table_body += append_total_row(total_other_income, "Other Income");
			// 	}
			// 	else if (message[1][i]["account"] == "Operating Expenses") {
			// 		table_body += append_total_row(total_other_income, "Other Income");
			// 	}
			// 	else if (message[1][i]["account"] == "Product Sales") {
			// 		table_body += append_total_row(total_other_income, "Other Income");
			// 	}
			// }

			
			
			table_body += append_group_row(account);
		} else {
			if (message[1][i]["account"] != "") {

				// new row with the gathered data
				if (parent == "Other Income - WW") 
					table_body += append_data_row(total_other_income, account, curr_data, prev_data, curr_ytd, prev_ytd);
				else if (parent == "Product Sales - WW") 
					table_body += append_data_row(total_product_sales, account, curr_data, prev_data, curr_ytd, prev_ytd);
				else if (parent == "Operating Expenses - WW") 
					table_body += append_data_row(total_operating_expenses, account, curr_data, prev_data, curr_ytd, prev_ytd);
				else if (parent == "Other Expenses - WW") 
					table_body += append_data_row(total_other_expenses, account, curr_data, prev_data, curr_ytd, prev_ytd);
			}
		}
		
		// this row needs to have a Gross Margin row that is not included in the gathered data
		if (account == "Cost of Goods Sold")
			table_body += append_gross_margin(total_income, account, curr_data, prev_data, curr_ytd, prev_ytd);

		// reset variables
		account    = "";
		curr_data  = "";
		prev_data  = "";
		curr_ytd   = "";
		prev_ytd   = "";
		is_group   = "";
		parent = ""
	}
	// table_body += append_total_row(total_expense);
	table_body += '</tbody>'; // end html table body

	return table_body;
}

function generate_table_css() {
	var table_css = "";

	table_css += '<style>';
	table_css += '    .table-data-right { font-family: Calibri; font-weight: normal; text-align: right; }';
	table_css += '    .table-data-left  { font-family: Calibri; font-weight: normal; text-align: left;  }';
	table_css += '</style>';

	return table_css;
}

function generate_table_caption(company, month, year) {
	var table_caption = "";

	table_caption += '<caption style="text-align: left;">';
	table_caption += '    <span style="font-family: Calibri; font-size: 10pt; text-align: left;">' + company + '</br></span>';
	table_caption += '    <span style="font-family: Calibri; font-size: 10pt; text-align: left;">' + month + ' 31, ' + year + '</span>';
	table_caption += '</caption>';

	return table_caption;
}

function generate_table_head(month, year) {
	var table_head = "";

	table_head += '<thead>';
	table_head += '    <tr style="border-top: 1px solid black; border-bottom: 1px solid black;">';
	table_head += '        <th style="text-align: right" colspan=3></th>';
	table_head += '        <th style="text-align: right" colspan=2>' + get_formatted_date(month, year, 0) + '</th>';
	table_head += '        <th style="text-align: right" colspan=2></th>';
	table_head += '        <th style="text-align: right" colspan=2>' + get_formatted_date(month, year, 1) + '</th>';
	table_head += '        <th style="text-align: right" colspan=2></th>';
	table_head += '        <th style="text-align: right" colspan=2>YTD-' + year.toString() + '</th>';
	table_head += '        <th style="text-align: right" colspan=2></th>';
	table_head += '        <th style="text-align: right" colspan=2>YTD-' + (parseInt(year) - 1).toString() + '</th>';
	table_head += '        <th style="text-align: right" colspan=2></th>';
	table_head += '    </tr>';
	table_head += '</thead>';

	return table_head;
}

// ======================================================================
// GETTER FUNCTIONS
// ======================================================================

function get_formatted_date(month, year, offset) {
	return (month.toString() + " " + (parseInt(year) - offset).toString());
}

function get_category_total(category_name, message, curr_month_year, prev_month_year) {
	let nf = new Intl.NumberFormat('en-US');
	var total_values  = [0.0, 0.0, 0.0, 0.0]; // array of totals for Income
	var index = 0;
	
	// find the beginning of this category and keep the index
	while (message[1][index]["account"] != category_name && index < message[1].length) index++;
	
	// check that it is a subgroup of either Income or Expense
	if (message[1][index]["indent"] == 1) index++;

	// everything under this subgroup is summed together into the array
	while (message[1][index]["indent"] == 2 && index < message[1].length) {
		total_values[0] += message[1][index][curr_month_year]
		total_values[1] += message[1][index][prev_month_year];
		total_values[2] += message[1][index]["total"];
		total_values[3] += message[1][index]["prev_year_total"];

		index++

		// break if end of array
		if (!message[1][index]) break;
	}

	// round down all the values before returning the array
	for (let j = 0; j < total_values.length; j++) {
		nf.format(Math.floor(total_values[j]));
	}
	
	return total_values;
}

// ======================================================================
// APPEND FUNCTIONS
// ======================================================================

function append_group_row(account) {
	var group_html = "";

	group_html += '<tr>';
	group_html += '    <td colspan=3>' + account + '</td>';
	group_html += '</tr>';

	return group_html
}

function append_data_row(total_array, account, curr_data, prev_data, curr_ytd, prev_ytd) {
	let nf = new Intl.NumberFormat('en-US');
	var data_html = "";

	var values = [
		(nf.format(Math.floor(curr_data))),
		(nf.format(Math.floor(prev_data))),
		(nf.format(Math.floor(curr_ytd))),
		(nf.format(Math.floor(prev_ytd))) 
	];

	var percentages = [
		((curr_data * 100) / total_array[0]).toString() + "%",
		((prev_data * 100) / total_array[1]).toString() + "%",
		((curr_ytd * 100)  / total_array[2]).toString() + "%",
		((prev_ytd * 100)  / total_array[3]).toString() + "%"
	];

	data_html += '<tr>';
	data_html += '    <td class="table-data-right" colspan=3>' + account + '</td>';
	for (let i = 0; i < 4; i++) {
		data_html += '<td class="table-data-right" colspan=2>' + values[i] + '</td>';
		data_html += '<td class="table-data-right" colspan=2>' + percentages[i] + '</td>';
	}
	data_html += '</tr>';

	return data_html;
}

function append_total_row(total_income, category_name) {
	let nf = new Intl.NumberFormat('en-US');
	var indent = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";  // prints an indent
	var total_row = "";

	var values = [
		nf.format(Math.floor(total_income[0])),
		nf.format(Math.floor(total_income[1])),
		nf.format(Math.floor(total_income[2])),
		nf.format(Math.floor(total_income[3]))
	]

	total_row += '<tr style="border-top: 1px solid black">';
	total_row += '    <td class="table-data-right" colspan=3>' + indent + 'Total ' + category_name + '</td>';
	for (let i = 0; i < 4; i++) {
		total_row += '<td class="table-data-right" colspan=2>' + values[i] + '</td>';
		total_row += '<td class="table-data-right" colspan=2>100%</td>';
	}
	total_row += '</tr>';
	total_row += '<tr></tr>';

	return total_row;
}

function append_gross_margin(total_income, account, curr_data, prev_data, curr_ytd, prev_ytd) {
	let nf = new Intl.NumberFormat('en-US');
	var gross_margin = "";
	
	var values = [
		nf.format(Math.floor(total_income[0] - curr_data)),
		nf.format(Math.floor(total_income[1] - prev_data)),
		nf.format(Math.floor(total_income[2] - curr_ytd)),
		nf.format(Math.floor(total_income[3] - prev_ytd))
	];

	var percentages = [
		((total_income[0] - curr_data)*100 / total_income[0]).toString() + "%",
		((total_income[1] - prev_data)*100 / total_income[1]).toString() + "%",
		((total_income[2] - curr_ytd)*100 / total_income[2]).toString()  + "%",
		((total_income[3] - prev_ytd)*100 / total_income[3]).toString()  + "%"
	];

	gross_margin += '<tr style="border-top: 1px solid black">';
	gross_margin += '    <td class="table-data-right" colspan=3>Gross Margin</td>';
	for (let i = 0; i < 4; i++) {
		data_html += '<td class="table-data-right" colspan=2>' + values[i] + '</td>';
		data_html += '<td class="table-data-right" colspan=2>' + percentages[i] + '</td>';
	}
	gross_margin += '</tr>';
	gross_margin += '<tr></tr>';

	return gross_margin;
}

// ======================================================================
// EXPORT FUNCTION
// ======================================================================

var tables_to_excel = (function () {
	var uri = 'data:application/vnd.ms-excel;base64,',
	html_start = `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">`,
	template_ExcelWorksheet = `<x:ExcelWorksheet><x:Name>{SheetName}</x:Name><x:WorksheetSource HRef="sheet{SheetIndex}.htm"/></x:ExcelWorksheet>`,
	template_ListWorksheet = `<o:File HRef="sheet{SheetIndex}.htm"/>`,
	template_HTMLWorksheet = `
------=_NextPart_dummy
Content-Location: sheet{SheetIndex}.htm
Content-Type: text/html; charset=windows-1252

` + html_start + `
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link id="Main-File" rel="Main-File" href="../WorkBook.htm">
<link rel="File-List" href="filelist.xml">
</head>
<body><table>{SheetContent}</table></body>
</html>`,
	template_WorkBook = `MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary="----=_NextPart_dummy"

------=_NextPart_dummy
Content-Location: WorkBook.htm
Content-Type: text/html; charset=windows-1252

` + html_start + `
<head>
<meta name="Excel Workbook Frameset">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="File-List" href="filelist.xml">
<!--[if gte mso 9]><xml>
<x:ExcelWorkbook>
<x:ExcelWorksheets>{ExcelWorksheets}</x:ExcelWorksheets>
<x:ActiveSheet>0</x:ActiveSheet>
</x:ExcelWorkbook>
</xml><![endif]-->
</head>
<frameset>
<frame src="sheet0.htm" name="frSheet">
<noframes><body><p>This page uses frames, but your browser does not support them.</p></body></noframes>
</frameset>
</html>
{HTMLWorksheets}
Content-Location: filelist.xml
Content-Type: text/xml; charset="utf-8"

<xml xmlns:o="urn:schemas-microsoft-com:office:office">
<o:MainFile HRef="../WorkBook.htm"/>
{ListWorksheets}
<o:File HRef="filelist.xml"/>
</xml>
------=_NextPart_dummy--
`,
	base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) },
	format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
	return function (tables, filename) {
		var context_WorkBook = {
			ExcelWorksheets: '',
			HTMLWorksheets: '',
			ListWorksheets: ''
		};

		var tables = jQuery(tables);
		var tbk = 0

		$.each(tables, function (SheetIndex, val) {
			var $table = $(val);
			var SheetName = "";
			
			if (SheetIndex == 0) {
				SheetName = 'IS - Consolidated';
			}
			
			context_WorkBook.ExcelWorksheets += format(template_ExcelWorksheet, {
				SheetIndex: SheetIndex,
				SheetName: SheetName
			});
			
			context_WorkBook.HTMLWorksheets += format(template_HTMLWorksheet, {
				SheetIndex: SheetIndex,
				SheetContent: $table.html()
			});
			
			context_WorkBook.ListWorksheets += format(template_ListWorksheet, {
				SheetIndex: SheetIndex
			});
			
			tbk += 1
		});

		var link = document.createElement("A");
		link.href = uri + base64(format(template_WorkBook, context_WorkBook));
		link.download = filename || 'Workbook.xls';
		link.target = '_blank';
		document.body.appendChild(link);
		link.click();
		document.body.removeChild(link);
	}
})();
