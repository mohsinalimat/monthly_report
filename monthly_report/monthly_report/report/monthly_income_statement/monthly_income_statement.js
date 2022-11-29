// ============================================================================================================================================
// INTRO
// ============================================================================================================================================
/* 
	To whoever is reading this, if this is the first time you are looking at this code, press Ctrl+K+0 to collapse everything on VS Code.
	
	It'll give you an outline of the functions I wrote and, in effect, a better understanding of the flow of the code. You can expand it
	back again by Ctrl+K+J.

	The functions are written in order of appearance to the best of my ability. frappe.call()'s callback function calls generate_tables(), so
	begin at generate_tables() and move along downwards. 

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

frappe.query_reports["Monthly Income Statement"] = {
	"filters": [
		{"fieldname": 'company',          "label": "Company",         "fieldtype": 'Link',   "options": 'Company', "default": frappe.defaults.get_user_default('company'), "hidden": true,},
		{"fieldname": "finance_book",     "label": "Finance Book",    "fieldtype": "Link",   "options": "Finance Book", "hidden": true,},
		{"fieldname": "to_fiscal_year",   "label": "End Year",        "fieldtype": "Link",   "options": "Fiscal Year", "default": frappe.defaults.get_user_default("fiscal_year")-1, "reqd": 1, "depends_on": "eval:doc.filter_based_on == 'Fiscal Year'"},
		{"fieldname": "period_end_month", "label": "Month",           "fieldtype": "Select", "options": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], "default": "January", "mandatory": 0, "wildcard_filter": 0},
		{"fieldname": "periodicity",      "label": "Periodicity",     "fieldtype": "Select", "options": [{ "value": "Monthly", "label": __("Monthly") }], "default": "Monthly", "reqd": true, "hidden": true,},
		{"fieldname": "filter_based_on",  "label": "Filter Based On", "fieldtype": "Select", "options": ["Fiscal Year", "Date Range"], "default": ["Fiscal Year"], "reqd": true, "hidden": true},
		{"fieldname": "cost_center",      "label": "Cost Center",     "fieldtype": "MultiSelectList", get_data: function (txt) {return frappe.db.get_link_options('Cost Center', txt, {company: frappe.query_report.get_filter_value("company")});}},
	],

	onload: function(report) {
		report.page.add_inner_button(__("Export Report"), function () {
			let filters = report.get_values();

			// stores the names of the cost centers from the filters for later use
			var cc_in_filters = filters.cost_center;

			// an array to store the consolidated, followed by each cost center's data
			var cc_consolidated = [];

			// names of all cost centers
			var cc_names = [
				'01 - White-Wood Corporate - WW',
				'02 - White-Wood Distributors Winnipeg - WW',
				'03 - Forest Products - WW',
				'06 - Endeavours - WW',
				'consolidated'
			];

			// loop through cost centers and call get_records() on each -- gather single cost center data
			// and append the dataset to cc_consolidated[]
			for (let i = 0; i < cc_names.length; i++) {
				if (cc_in_filters.includes(cc_names[i])) {
					filters.cost_center = [cc_names[i]]; 

					frappe.call({
						method: 'monthly_report.monthly_report.report.monthly_income_statement.monthly_income_statement.get_records',
						args: {filters: filters},
		
						callback: function (r) {
							cc_consolidated.push(r.message);
							show_alert('Gathered Data for ' + cc_names[i].slice(0,-5), 5);
						}
					});

				} else if (cc_names[i] == "consolidated") {
					filters.cost_center = cc_in_filters; // restore the original list of cost centers

					frappe.call({
						method: 'monthly_report.monthly_report.report.monthly_income_statement.monthly_income_statement.get_records',
						args: {filters: filters},

						callback: function (r) {
							cc_consolidated.push(r.message); // push the consolidated dataset to the end
							cc_consolidated.reverse(); // reverse the order of the array to bring consolidated to [0]

							// finally generate the tables using each cost center dataset
							generate_tables(cc_consolidated, filters.company, filters.period_end_month, filters.to_fiscal_year, filters.cost_center, true)
						}
					});
				}
				wait(1000);
			}
		});
	},
}

var wait = (ms) => {
    const start = Date.now();
    let now = start;
    while (now - start < ms) {
      now = Date.now();
    }
}

// ============================================================================================================================================
// GLOBAL FLAGS
// ============================================================================================================================================

var minus_to_brackets = 0; 	// flag that determines if negative numbers are to be represented with brackets instead: i.e., -1 to (1)
var capitalized_names = 0; 	// account names will be in block letters or sentence case
var download_excel = 1; 	// flag that determines if the excel spreadsheet is to be downloaded at the end of processing
var console_log = 0; 		// flag that determines if console logs should be printed

// ============================================================================================================================================
// GENERATOR FUNCTIONS
// ============================================================================================================================================

function generate_tables(dataset, company, month, year, cost_centers, consolidate) {
	// dataset[0] contains the consolidated data
	message = dataset[0];

	// give each table an ID to later identify them in the export function
	var $table_id  = "consolidated";
	var tables_array = [("#" + $table_id)];

	// date info needed to generate the tables
	var month_name = (month.slice(0, 3)).toLowerCase();
	var curr_month_year = month_name + "_" + year;
	var prev_month_year = month_name + "_" + (parseInt(year) - 1).toString();

	// html that gets sent to the export function
	// populate it with consolidated data as it must always be present
	var html = generate_consolidated(company, month, year, message, curr_month_year, prev_month_year, $table_id, "Consolidated Income Statement");

	if (cost_centers.length == 1) { // if only one cost center was chosen we just populated both spreadsheets with the same data as they are identical
		$table_id  = "table_0";
		tables_array.push("#" + $table_id);
		html += generate_consolidated(company, month, year, message, curr_month_year, prev_month_year, $table_id, (cost_centers[0].slice(5, -5) + " Income Statement"));
	
	} else { // when there are multiple cost centers we process them differently
		for (let i = 0; i < cost_centers.length; i++) {
			$table_id  = "table_" + i;
			tables_array.push("#" + $table_id);
	
			html += '<div id="data">';
			html += '<table style="font-weight: normal; font-family: Calibri; font-size: 10pt" id=' + $table_id + '>';
			html += generate_table_caption(company, month, year, (cost_centers[i].slice(5, -5) + " Income Statement"));
			html += generate_table_head(month, year);
			html += generate_table_body(dataset[i+1], curr_month_year, prev_month_year);
			html += '</table>';
			html += '</div>';
		}	
	}

	// append the css & html, then export as xls
	$(".report-wrapper").hide();
	$(".report-wrapper").append(html);

	// process the cost center numbers to creat the excel sheet names on the tabs
	var center_numbers = ["0"];
	for (let i = 0; i < cost_centers.length; i++)
		center_numbers.push(cost_centers[i].slice(1, 2));

	// flags to control the export and download -- used for testing without filling the downloads folder with junk
	if (download_excel && consolidate)
		tables_to_excel(tables_array, curr_month_year +'.xls', center_numbers);
}

function generate_consolidated(company, month, year, message, curr_month_year, prev_month_year, $table_id, title) {
	// css for the table 
	var html = generate_table_css();
	
	// the table containing all the data in html format
	html += '<div id="data">';
	html += '<table style="font-weight: normal; font-family: Calibri; font-size: 10pt" id=' + $table_id + '>';
	html += generate_table_caption(company, month, year, title);
	html += generate_table_head(month, year);
	html += generate_table_body(message, curr_month_year, prev_month_year);
	html += '</table>';
	html += '</div>';

	return html;
}

function generate_table_css() {
	var table_css = "";

	table_css += '<style>';
	table_css += '    .table-data-right { font-family: Calibri; font-weight: normal; text-align: right; }';
	table_css += '    .table-data-left  { font-family: Calibri; font-weight: normal; text-align: left;  }';
	table_css += '</style>';

	return table_css;
}

function generate_table_caption(company, month, year, title) {
	var table_caption = "";

	table_caption += '<caption style="text-align: left;">';
	table_caption += '    <span style="font-family: Calibri; font-size: 10pt; text-align: left;">' + company + '</br></span>';
	table_caption += '    <span style="font-family: Calibri; font-size: 10pt; text-align: left;">' + title + '</br></span>';
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

	var categories = [
		"Product Sales",
		"Other Income",
		"Cost of Goods Sold",
		"Operating Expenses",
		"Other Expenses"
	]
	
	html += '<tbody>'; // start html table body
	for (let i = 0; i < categories.length; i++) {
		if (i == 0)
			html += append_group_row("Income");
		if (i == 3) 
			html += append_group_row("Expenses");

		if (category_exists(message, categories[i]))
			html += get_category_rows(categories[i], message, curr_month_year, prev_month_year);
		else 
			if (console_log) 
				console.log("[!] " + categories[i] + " does not exists");
	}
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
			let start = account_object["account"].indexOf("-");
			account += account_object["account"].slice(start+2);
			
		} else {
			account += account_object["account"];
	} else {
		let start = account_object["account_name"].indexOf("-");
		account += account_object["account_name"].slice(start+2);
	}

	// corrects account names in  block letters
	if (account.includes("-")) {
		// split the word by "-"
		var split_words = account_object["account"].split("-");
		
		// keep the first letter of each word capital, and append the rest in lowercase
		for (let i = 1; i < split_words.length; i++) {
			split_words[i] = split_words[i].trim();
			let temp_word = split_words[i][0];
			temp_word += split_words[i].slice(1).toLowerCase();
			split_words[i] = temp_word;
		}

		// clear the variable and add indentation
		account = "";
		for (let i = 0; i < account_object["indent"]; i++)
			account += indent;

		// append the case corrected words and hyphens
		for (let i = 1; i < split_words.length; i++){
			account += split_words[i];
			if (i != split_words.length-1) 
				account += "-";
		}
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
	var print_groups = [];

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

	index++;

	// gather each row's data under the current category
	// also compresses the accounts based on print groups

	// if a print group already exists in print_groups[], sum up their values 
	// if it does not exist already, append that group to print_groups[] along with its data
	while (message[1][index]["parent_account"] == category_name) {
		account   = message[1][index]["account"];
		curr_data = message[1][index][curr_month_year];
		prev_data = message[1][index][prev_month_year];
		curr_ytd  = message[1][index]["total"];
		prev_ytd  = message[1][index]["prev_year_total"];
		indent    = message[1][index]["indent"];
		is_group  = message[1][index]["is_group"];

		// this section compares the current print group against existing print groups
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

		// if the print group was not found, append it to the end
		if (!group_found) {
			// create an object containing the info and push() to print_groups[]
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
		
		// break the loop if no more rows exist in the source array
		if (!message[1][index]) 
			break;
	}
	
	// adds each row's gathered data to the html
	for (let i = 0; i < print_groups.length; i++) {
		account   = get_formatted_name(print_groups[i]);
		curr_data = print_groups[i].curr_data;
		prev_data = print_groups[i].prev_data;
		curr_ytd  = print_groups[i].curr_ytd;
		prev_ytd  = print_groups[i].prev_ytd;

		html += append_data_row(category_total, account, curr_data, prev_data, curr_ytd, prev_ytd);
	}

	// appends a row containing the total values for the current category
	html += append_total_row(category_name, category_total);

	if (console_log) console.log(" --> Appended " + category_name);

	return html;
}

function get_category_total(category_name, message, curr_month_year, prev_month_year) {
	if (console_log) console.log("\t[" + category_name + "] calculating total");

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
	
	if (console_log) console.log("\t[" + category_name + "] total calculated");
	return total_values;
}

function category_exists(message, category_name) {
	var category_exists = false;

	for (let i = 0; i < message[1].length; i++) {
		if (message[1][i]["account"] == category_name){
			category_exists = true;
			break;
		}
	}

	if (category_exists) 
		if (console_log) console.log(category_name + " exists");

	return category_exists;
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

		if (percentages[i].toString().slice(0, -1) == "NaN" || percentages[i].toString().slice(0, -1) == "100.00") 
			html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>100%</td>';
		else if (percentages[i].toString().slice(0, -1) == "0.00")
			html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>0%</td>';
		else
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
	if (console_log) console.log("Calculating Gross Margin");

	let nf = new Intl.NumberFormat('en-US');
	var html = "";
	var total_income = [0.0, 0.0, 0.0, 0.0];
	var total_pd   = [0.0, 0.0, 0.0, 0.0];
	var total_oi   = [0.0, 0.0, 0.0, 0.0];
	var total_cogs = [0.0, 0.0, 0.0, 0.0];
	
	if (category_exists(message, "Product Sales"))
		total_pd = get_category_total("Product Sales", message, curr_month_year, prev_month_year);

	if (category_exists(message, "Other Income"))
		total_oi = get_category_total("Other Income", message, curr_month_year, prev_month_year);

	total_cogs = get_category_total("Cost of Goods Sold", message, curr_month_year, prev_month_year);

	for (let i = 0; i < 4; i++)
		total_income[i] += (total_pd[i] + total_oi[i]);

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
		html += '<td class="table-data-right" style="font-size: 10pt; text-align: right;" colspan=1>' + total_cogs[i] + '</td>';
		html += '<td class="table-data-right" style="font-size: 10pt; text-align: right;" colspan=1>' + cogs_percentages[i] + '</td>';
	}
	html += '</tr>';

	html += '<tr style="border-top: 1px solid black">';
	html += '<td class="table-data-right" style="font-size: 10pt" colspan=3>Gross Margin</td>';
	for (let i = 0; i < 4; i++) {
		html += '<td class="table-data-right" style="font-size: 10pt; text-align: right;" colspan=1>' + values[i] + '</td>';
		html += '<td class="table-data-right" style="font-size: 10pt; text-align: right;" colspan=1>' + percentages[i] + '</td>';
	}
	html += '</tr>';
	html += '<tr></tr>';

	if (console_log) console.log("Appended Gross Margin");
	return html;
}

// ============================================================================================================================================
// EXPORT FUNCTION
// ============================================================================================================================================

var tables_to_excel = (function () {
	var uri = 'data:application/vnd.ms-excel;base64,',
	
	html_start = (
		`<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">`
	),
	
	template_ExcelWorksheet = (
		`<x:ExcelWorksheet><x:Name>{SheetName}</x:Name><x:WorksheetSource HRef="sheet{SheetIndex}.htm"/></x:ExcelWorksheet>`
	),
	
	template_HTMLWorksheet = (`
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
</html>`
	),

	template_WorkBook = (`MIME-Version: 1.0
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
`
	),

	base64 = function (s) { 
		return window.btoa(unescape(encodeURIComponent(s))) 
	},

	format = function (s, c) {
		return s.replace(/{(\w+)}/g, function (m, p) { 
			return c[p]; 
		}) 
	}
	
	return function (tables, filename, center_numbers) {

		var context_WorkBook = {
			ExcelWorksheets: '',
			HTMLWorksheets: '',
		};

		var tables = jQuery(tables);

		$.each(tables, function (SheetIndex, val) {
			var $table = $(val);
			var SheetName = "";
			let center_number = center_numbers[SheetIndex];
			
			if (center_number == "0") 
				SheetName = "IS - Consolidated";
			else if (center_number == "1") 
				SheetName = "IS - White-Wood Corporate";
			else if (center_number == "2") 
				SheetName = "IS - White-Wood Distributors";
			else if (center_number == "3") 
				SheetName = "IS - Forest Products";
			else if (center_number == "6") 
				SheetName = "IS - Endeavours";
			else 
				SheetName = "-";
			
			context_WorkBook.ExcelWorksheets += format(template_ExcelWorksheet, {
				SheetIndex: SheetIndex,
				SheetName: SheetName
			});
			
			context_WorkBook.HTMLWorksheets += format(template_HTMLWorksheet, {
				SheetIndex: SheetIndex,
				SheetContent: $table.html()
			});
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
