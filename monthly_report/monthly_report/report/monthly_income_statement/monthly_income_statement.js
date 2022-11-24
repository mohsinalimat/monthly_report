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

// ============================================================================================================================================
// INTRO
// ============================================================================================================================================
/* 
	To whoever is reading this, if this is the first time you are looking at this code, press Ctrl+K+0 to collapse everything on VS Code.
	
	It'll give you an outline of the functions I wrote and, in effect, a better understanding of the flow of the code. You can expand it
	back again by Ctrl+K+J.

	The functions are written in order of appearance to the best of my ability. Begin at generate_table() and move along downwards. 

	To summarize, 
		-> the generator functions are the ones that generate the html that gets exported to excel.
		-> the getter functions are supporting functions that get you the thing it says in its name.
		-> the append functions are used to append one row or section of html based on the data passed to it.
		-> tables_to_excel() is some sort of legacy code that I did not write, and is difficult to read. I've left it mostly untouched,
			but it does what it describes -- converts the tables to excel.
		-> the global flags are essentially shorcuts to avoid commenting out stuff when I needed to test things.

	Everything is made to be modular, as best as I could. Moving things around a bit should still allow things to work.

	Hope this helps.

	- Farabi Hussain
*/ 

// ============================================================================================================================================
// GLOBAL FLAGS
// ============================================================================================================================================

var minus_to_brackets = false; // flag that determines if negative numbers are to be represented with brackets instead: i.e., -1 to (1)
var download_excel = true;     // flag that determines if the excel spreadsheet is to be downloaded at the end of processing

// ============================================================================================================================================
// GENERATOR FUNCTIONS
// ============================================================================================================================================

function generate_table(message, company, month, year) {
	// tbh i dont know what these are, just that things break without them 
	var $export_id  = "export_id";
	var table_count = [];
	table_count[0]  = "#" + $export_id;

	var table_html = "";
	var table_css  = "";
	var month_name = (month.slice(0, 3)).toLowerCase();
	var curr_month_year = month_name + "_" + year;
	var prev_month_year = month_name + "_" + (parseInt(year) - 1).toString();

	// css for the table 
	table_css = generate_table_css();
	
	// the table containing all the data in html format
	table_html += '<div id="data">';
	table_html += '<table style="font-weight: normal; font-family: Calibri; font-size: 10pt" id=' + $export_id + '>';
	table_html += generate_table_caption(company, month, year)
	table_html += generate_table_head(month, year)
	table_html += generate_table_body(message, curr_month_year, prev_month_year)
	table_html += '</table>';
	table_html += '</div>';

	// append the css & html, then export as xls
	$(".report-wrapper").hide();
	$(".report-wrapper").append(table_css);
	$(".report-wrapper").append(table_html);

	if (download_excel)
		tables_to_excel(table_count, curr_month_year +'.xls');
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
	table_caption += '    <span style="font-family: Calibri; font-size: 10pt; text-align: left;">' + month + '&nbsp;31,&nbsp;' + year + '</span>';
	table_caption += '</caption>';

	return table_caption;
}

function generate_table_head(month, year) {
	var table_head = "";

	table_head += '<thead>';
	table_head += '    <tr style="border-top: 1px solid black; border-bottom: 1px solid black;">';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=3></th>';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=1>' + get_formatted_date(month, year, 0) + '</th>';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=1></th>';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=1>' + get_formatted_date(month, year, 1) + '</th>';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=1></th>';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=1>YTD ' + year.toString() + '</th>';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=1></th>';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=1>YTD ' + (parseInt(year) - 1).toString() + '</th>';
	table_head += '        <th style="text-align: right; font-size: 10pt" colspan=1></th>';
	table_head += '    </tr>';
	table_head += '</thead>';

	return table_head;
}

function generate_table_body(message, curr_month_year, prev_month_year) {
	var html = ""; // holds the html that is returned

	html += '<tbody>'; // start html table body

	// adds the Income section: contains 'Product Sales' and 'Other Income'
	html += append_group_row("Income")
	html += get_category_rows("Product Sales", message, curr_month_year, prev_month_year);
	html += get_category_rows("Other Income", message, curr_month_year, prev_month_year);

	// adds the 'Total Cost of Goods' and the 'Gross Margin' rows
	html += append_gross_margin(message, curr_month_year, prev_month_year);

	// adds the Income section: contains 'Operating Expenses' and 'Other Expenses'
	html += append_group_row("Expenses")
	html += get_category_rows("Operating Expenses", message, curr_month_year, prev_month_year);
	html += get_category_rows("Other Expenses", message, curr_month_year, prev_month_year);
	html += '</tbody>'; // end html table body

	return html;
}

// ============================================================================================================================================
// GETTER FUNCTIONS
// ============================================================================================================================================

function get_formatted_name(account_object) {
	var account = "";
	var indent = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";  // prints an indent


	// add the indent in a loop
	for (let i = 0; i < account_object["indent"]; i++)
		account += indent;
	
	// check if the account field has the print group name, otherwise we print the account name
	// the split removes the number in front of the strings like "30 - Trade Sales" to "Trade Sales"
	if (account_object["account"] != "")
	 	if (account_object["is_group"] == 0) {
			console.log(account_object);

			let start = account_object["account"].indexOf("-");
			account += account_object["account"].slice(start+2);
		} else {
			account += account_object["account"];
	} else {
		let start = account_object["account_name"].indexOf("-");
		account += account_object["account_name"].slice(start+2);
	}

	return account;
}

function get_formatted_date(month, year, offset) {
	return ("&nbsp;" + month.toString().slice(0, 3) + " " + (parseInt(year) - offset).toString());
}

function get_formatted_number(number) {
	var formatted_number = "";

	// check for minus sign and add brackets if found
	if (number.toString()[0] == "-")
		formatted_number = "&nbsp;(" + number.toString().slice(1) + ")";
	else
		formatted_number = number.toString();


	if (minus_to_brackets) 
		return formatted_number;
	else 
		return number;
}

function get_category_rows(category_name, message, curr_month_year, prev_month_year) {
	var index = 0;
	var html = "";
	var category_total = get_category_total(category_name, message, curr_month_year, prev_month_year);

	let account   = "";
	let curr_data = "";
	let prev_data = "";
	let curr_ytd  = "";
	let prev_ytd  = "";
	let indent    = "";
	let is_group  = "";
	
	// find the beginning of this category and keep the index
	while (message[1][index]["account"] != category_name && index < message[1].length) 
		index++;

	html += append_group_row(get_formatted_name(message[1][index]));

	var print_groups = [];

	index++;
	while (message[1][index]["parent_account"] == category_name) {
		account   = message[1][index]["account"];
		curr_data = message[1][index][curr_month_year];
		prev_data = message[1][index][prev_month_year];
		curr_ytd  = message[1][index]["total"];
		prev_ytd  = message[1][index]["prev_year_total"];
		indent    = message[1][index]["indent"];
		is_group  = message[1][index]["is_group"];

		let group_found = false;
		for (let j = 0; j < print_groups.length; j++) {
			if (print_groups[j].account == account) {
				print_groups[j].curr_data += curr_data;
				print_groups[j].prev_data += prev_data;
				print_groups[j].curr_ytd  += curr_ytd;
				print_groups[j].prev_ytd  += prev_ytd;
				print_groups[j].indent    = indent;
				print_groups[j].is_group  = is_group;

				group_found = true;
				break;
			}
		}

		if (!group_found) {
			print_groups.push({
				"account"   : account,
				"curr_data" : curr_data,
				"prev_data" : prev_data,
				"curr_ytd"  : curr_ytd,
				"prev_ytd"  : prev_ytd,
				"indent"    : indent,
				"is_group"  : is_group
			});
		}
		
		index++;
		
		if (!message[1][index]) 
		break;
	}
	
	for (let i = 0; i < print_groups.length; i++) {
		account   = get_formatted_name(print_groups[i]);
		// account   = print_groups[i].account;
		curr_data = print_groups[i].curr_data;
		prev_data = print_groups[i].prev_data;
		curr_ytd  = print_groups[i].curr_ytd;
		prev_ytd  = print_groups[i].prev_ytd;

		html += append_data_row(category_total, account, curr_data, prev_data, curr_ytd, prev_ytd);
	}

	html += append_total_row(category_name, category_total);

	return html;
}

function get_category_total(category_name, message, curr_month_year, prev_month_year) {
	let nf = new Intl.NumberFormat('en-US');
	var total_values  = [0.0, 0.0, 0.0, 0.0]; // array of totals for Income
	var index = 0;
	
	// find the beginning of this category and keep the index
	while (message[1][index]["account"] != category_name && index < message[1].length) 
		index++;
	
	// check that it is a subgroup of either Income or Expense
	if (message[1][index]["indent"] == 1) 
		index++;

	// everything under this subgroup is summed together into the array
	while (message[1][index]["indent"] == 2 && index < message[1].length) {
		total_values[0] += message[1][index][curr_month_year]
		total_values[1] += message[1][index][prev_month_year];
		total_values[2] += message[1][index]["total"];
		total_values[3] += message[1][index]["prev_year_total"];

		index++

		// break if end of array
		if (!message[1][index])
			break;
	}

	// round down all the values before returning the array
	for (let j = 0; j < total_values.length; j++)
		nf.format(Math.floor(total_values[j]));
	
	return total_values;
}

// ============================================================================================================================================
// APPEND FUNCTIONS
// ============================================================================================================================================

function append_group_row(account) {
	var html = "";

	html += '<tr>';
	html += '    <td style="font-size: 10pt" colspan=3>' + account + '</td>';
	html += '</tr>';

	return html
}

function append_data_row(total_array, account, curr_data, prev_data, curr_ytd, prev_ytd) {
	let nf = new Intl.NumberFormat('en-US');
	var html = "";

	var values = [
		(nf.format(Math.floor(curr_data))),
		(nf.format(Math.floor(prev_data))),
		(nf.format(Math.floor(curr_ytd))),
		(nf.format(Math.floor(prev_ytd))) 
	];

	var percentages = [
		get_formatted_number(((curr_data * 100) / total_array[0]).toFixed(2)) + "%",
		get_formatted_number(((prev_data * 100) / total_array[1]).toFixed(2)) + "%",
		get_formatted_number(((curr_ytd * 100)  / total_array[2]).toFixed(2)) + "%",
		get_formatted_number(((prev_ytd * 100)  / total_array[3]).toFixed(2)) + "%"
	];

	html += '<tr>';
	html += '<td class="table-data-right" style="font-size: 10pt" colspan=3>' + account + '</td>';
	for (let i = 0; i < 4; i++) {
		html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>' + get_formatted_number(values[i]) + '</td>';
		html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>' + percentages[i] + '</td>';
	}
	html += '</tr>';

	return html;
}

function append_total_row(category_name, category_total) { 
	let nf = new Intl.NumberFormat('en-US');
	var indent = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";  // prints an indent
	var html = "";

	var values = [
		nf.format(Math.floor(category_total[0])),
		nf.format(Math.floor(category_total[1])),
		nf.format(Math.floor(category_total[2])),
		nf.format(Math.floor(category_total[3]))
	]

	html += '<tr style="border-top: 1px solid black">';
	html += '<td class="table-data-right" style="font-size: 10pt" colspan=3>' + indent + 'Total ' + category_name + '</td>';
	for (let i = 0; i < 4; i++) {
		html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + values[i] + '</td>';
		html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>100%</td>';
	}
	html += '</tr>';
	html += '<tr></tr>';

	return html;
}

function append_gross_margin(message, curr_month_year, prev_month_year) {
	let nf = new Intl.NumberFormat('en-US');
	var html = "";
	var total_income = [0.0, 0.0, 0.0, 0.0];

	var total_cogs  = get_category_total("Cost of Goods Sold", message, curr_month_year, prev_month_year);
	var total_sales = get_category_total("Product Sales", message, curr_month_year, prev_month_year);
	var total_other = get_category_total("Other Income", message, curr_month_year, prev_month_year);

	for (let i = 0; i < 4; i++)
		total_income[i] += (total_sales[i] + total_other[i]);

	var values = [
		total_income[0] - total_cogs[0],
		total_income[1] - total_cogs[1],
		total_income[2] - total_cogs[2],
		total_income[3] - total_cogs[3]
	];

	var percentages = [
		((total_income[0] - total_cogs[0])*100 / total_income[0]),
		((total_income[1] - total_cogs[1])*100 / total_income[1]),
		((total_income[2] - total_cogs[2])*100 / total_income[2]),
		((total_income[3] - total_cogs[3])*100 / total_income[3])
	];

	var cogs_percentages = [
		(100 - percentages[0]),
		(100 - percentages[1]),
		(100 - percentages[2]),
		(100 - percentages[3])
	];

	for (let i = 0; i < 4; i++) {
		total_cogs[i]       = nf.format(Math.floor(total_cogs[i])).toString();
		values[i]           = nf.format(Math.floor(values[i])).toString();
		percentages[i]      = percentages[i].toString() + "%";
		cogs_percentages[i] = cogs_percentages[i].toString() + "%";
	}

	html += '<tr>';
	html += '<td class="table-data-right" style="font-size: 10pt" colspan=3>Cost of Goods Sold</td>';
	for (let i = 0; i < 4; i++) {
		html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + total_cogs[i] + '</td>';
		html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + cogs_percentages[i] + '</td>';
	}
	html += '</tr>';

	html += '<tr style="border-top: 1px solid black">';
	html += '<td class="table-data-right" style="font-size: 10pt" colspan=3>Gross Margin</td>';
	for (let i = 0; i < 4; i++) {
		html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + values[i] + '</td>';
		html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + percentages[i] + '</td>';
	}
	html += '</tr>';
	html += '<tr></tr>';

	return html;
}

// ============================================================================================================================================
// EXPORT FUNCTION
// ============================================================================================================================================

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
