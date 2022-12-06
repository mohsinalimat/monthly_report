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
// GLOBAL FLAGS
// ============================================================================================================================================


var minus_to_brackets = 0; 	 // flag that determines if negative numbers are to be represented with brackets instead: i.e., -1 to (1)
var capitalized_names = 0; 	 // account names will be in block letters or sentence case
var download_excel = 1; 	 // flag that determines if the excel spreadsheet is to be downloaded at the end of processing
var console_log = 0; 		 // flag that determines if console logs should be printed
var report_type = "Regular"; // determines the type of report that will be generated

// ============================================================================================================================================
// GLOBAL VARIABLES
// ============================================================================================================================================


var indent = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"; // prints an indent of 8 spaces

var cc_names = [
    '01 - White-Wood Corporate - WW',
    '02 - White-Wood Distributors Winnipeg - WW',
    '03 - Forest Products - WW',
    '06 - Endeavours - WW',
];

var wait = (ms) => {
    const start = Date.now();
    let now = start;

    while (now - start < ms)
        now = Date.now();
}

frappe.query_reports["Monthly Income Statement"] = {
    "filters": [
        {"fieldname": 'company',          "label": "Company",         "fieldtype": 'Link',   "reqd": false, "hidden": true,  "default": frappe.defaults.get_user_default('company'),     "options": 'Company'},
        {"fieldname": "finance_book",     "label": "Finance Book",    "fieldtype": "Link",   "reqd": false, "hidden": true,                                                              "options": "Finance Book"},
        {"fieldname": "to_fiscal_year",   "label": "End Year",        "fieldtype": "Link",   "reqd": true,  "hidden": false, "default": frappe.defaults.get_user_default("fiscal_year"), "options": "Fiscal Year", "depends_on": "eval:doc.filter_based_on == 'Fiscal Year'"},
        {"fieldname": "period_end_month", "label": "Month",           "fieldtype": "Select", "reqd": true,  "hidden": false, "default": "January", "mandatory": 0, "wildcard_filter": 0, "options": ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]},
        {"fieldname": "periodicity",      "label": "Periodicity",     "fieldtype": "Select", "reqd": true,  "hidden": true,  "default": "Monthly",                                       "options": [{ "value": "Monthly", "label": __("Monthly") }]},
        {"fieldname": "filter_based_on",  "label": "Filter Based On", "fieldtype": "Select", "reqd": true,  "hidden": true,  "default": ["Fiscal Year"],                                 "options": ["Fiscal Year", "Date Range"]},
        {"fieldname": "report_type",      "label": "Report Type",     "fieldtype": "Select", "reqd": true,  "hidden": false, "default": ["Regular"],                                     "options": ["Expanded", "Regular", "Minimal"]},
        {"fieldname": "cost_center",      "label": "Cost Center",     "fieldtype": "MultiSelectList", get_data: function (txt) {return frappe.db.get_link_options('Cost Center', txt, {company: frappe.query_report.get_filter_value("company")});}},
    ],

    onload: function(report) {
        report.page.add_inner_button(__("Export Report"), function () {
            let filters = report.get_values();

            /*             
            // ------------- for testing, make sure this is commented out -------------
            filters.cost_center = [
                '01 - White-Wood Corporate - WW',
                '02 - White-Wood Distributors Winnipeg - WW',
                '03 - Forest Products - WW',
                '06 - Endeavours - WW',
            ];
            // ------------- for testing, make sure this is commented out -------------
            */

            // an array to store the consolidated, followed by each cost center's data
            var dataset = [];

            show_alert({message: 'Retrieving data for ' + filters.cost_center.length + ' cost centers', indicator: 'blue'}, (filters.cost_center.length*15));

            // retrieve the consolidated dataset for all selected cost centers
            frappe.call({
                method: 'monthly_report.monthly_report.report.monthly_income_statement.custom_monthly_report_script.generate_monthly_report',
                args: {filters: filters},

                callback: function (r) {
                    dataset.push(r.message); // push the consolidated dataset to the end
                    
                    show_alert({message: 'Exporting spreadsheet', indicator: 'green'}, 4);
                    // finally generate the tables using each cost center dataset
                    generate_tables(dataset, filters.company, filters.period_end_month, filters.to_fiscal_year, filters.cost_center);
                }
            });
        });
    },
}


// ============================================================================================================================================
// GENERATOR FUNCTIONS
// ============================================================================================================================================


// generates the entire table by calling functions that generate the css, caption, header, and body
function generate_tables(dataset, company, month, year, cost_centers) {
    // dataset[0] contains the consolidated data
    consolidated_data = dataset[0][1];
    balance_sheet_data = dataset[0][3];

    var html = "";
    var $table_id = "";
    var tables_array = [];

    // give each table an ID to later identify them in the export function
    var $table_id = "consolidated";
    var tables_array = [("#" + $table_id)];

    // date info needed to generate the tables
    var month_name = (month.slice(0, 3)).toLowerCase();
    var curr_month_year = month_name + "_" + year;
    var prev_month_year = month_name + "_" + (parseInt(year) - 1).toString();

    // only one cost center was chosen we just populated both spreadsheets with the same data as they are identical
    if (cost_centers.length == 1) { 
        // give each table an ID to later identify them in the export function
        html += generate_consolidated(company, month, year, consolidated_data, curr_month_year, prev_month_year, $table_id, (cost_centers[0].slice(5, -5) + " Income Statement"), mode = "income_statement");

        $table_id = "table_0";
        tables_array.push("#" + $table_id);
        html += generate_consolidated(company, month, year, consolidated_data, curr_month_year, prev_month_year, $table_id, (cost_centers[0].slice(5, -5) + " Income Statement"), mode = "trailing_12_months");
    
        // appends the Balance Sheet
        $table_id = "balance_sheet";
        tables_array.push("#" + $table_id);
        html += generate_balance_sheet(company, month, year, balance_sheet_data, curr_month_year, prev_month_year, $table_id, mode = "balance_sheet");

    // when there are multiple cost centers we process them differently
    } else { 
        // give each table an ID to later identify them in the export function
        $table_id = "consolidated";
        tables_array = [("#" + $table_id)];

        // populate it first with consolidated data as it must always be present
        html = generate_consolidated(company, month, year, consolidated_data, curr_month_year, prev_month_year, $table_id, "Consolidated Income Statement", mode = "income_statement");

        // appends individual cost center income statement sheets
        let id = 0;

        for (let i = 0; i < cost_centers.length; i++) {
            $table_id = "table_" + id;
            tables_array.push("#" + $table_id);
    
            html += '<div id="data">';
            html += '<table style="font-weight: normal; font-family: Calibri; font-size: 10pt" id=' + $table_id + '>';
            html += generate_table_caption(company, month, year, (cost_centers[i].slice(5, -5) + " Income Statement"));
            html += generate_table_head(month, year, mode = "income_statement");

            // [0][j][0]["dataset_for"] contains the cost center name 
            // I used that to cross check and populate the sheet so that none of the data gets swapped across different tabs
            // i.e., tab named Endeavours might contain Forest's data otherwirse since Javascript is asynchronous
            for (let j = 0; j < dataset[0].length; j++)
                if (dataset[0][j][0]["dataset_for"] === cost_centers[i])
                    html += generate_table_body(dataset[0][j+1], curr_month_year, prev_month_year, mode = "income_statement");

            html += '</table>';
            html += '</div>';
            id++;
        }


        // appends the Trailing Twelve Months sheets
        id++;
        $table_id = "table_" + id;
        tables_array.push("#" + $table_id);
        html += generate_consolidated(company, month, year, consolidated_data, curr_month_year, prev_month_year, $table_id, "Consolidated Income Statement", mode = "trailing_12_months");

        id++;
        for (let i = 0; i < cost_centers.length; i++) {
            $table_id = "table_" + id;
            tables_array.push("#" + $table_id);

            html += '<div id="data">';
            html += '<table style="font-weight: normal; font-family: Calibri; font-size: 10pt" id=' + $table_id + '>';
            html += generate_table_caption(company, month, year, (cost_centers[i].slice(5, -5) + " Income Statement"));
            html += generate_table_head(month, year, mode);

            for (let j = 1; j < dataset[0].length; j++)
                if (dataset[0][j][0]["dataset_for"] === cost_centers[i])
                    html += generate_table_body(dataset[0][j+1], curr_month_year, prev_month_year, mode = "trailing_12_months");

            html += '</table>';
            html += '</div>';
            id++;
        }

        
        // appends the Balance Sheet
        $table_id = "balance_sheet";
        tables_array.push("#" + $table_id);
        mode = "balance_sheet";
        html += generate_balance_sheet(company, month, year, balance_sheet_data, curr_month_year, prev_month_year, $table_id, mode);
    }

    // append the css & html, then export as xls
    $(".report-wrapper").hide();
    $(".report-wrapper").append(html);

    // process the cost center numbers to creat the excel sheet names on the tabs
    var center_numbers = [];

    // append the *Income Statement* numbers to the list of cost centers
    if (cost_centers.length > 1) 
        center_numbers.push("0_IS");

    for (let i = 0; i < cost_centers.length; i++)
        center_numbers.push(cost_centers[i].slice(1, 2) + "_IS");

    // append the *Trailing 12 Months* numbers to the list of cost centers
    if (cost_centers.length > 1) 
        center_numbers.push("0_TTM");

    for (let i = 0; i < cost_centers.length; i++)
        center_numbers.push(cost_centers[i].slice(1, 2) + "_TTM");

    // append the *Balance Sheet* number
    center_numbers.push("0_BS");

    console.log(center_numbers);

    // flags to control the export and download -- used for testing without filling the downloads folder with junk
    if (download_excel)
        tables_to_excel(tables_array, curr_month_year + '.xls', center_numbers);
}

// shortcut that generates the consolidated table without extra adjustments 
function generate_consolidated(company, month, year, dataset, curr_month_year, prev_month_year, $table_id, title, mode) {
    // css for the table 
    var html = generate_table_css();
    
    // the table containing all the data in html format
    html += '<div id="data">';
    html += '<table style="font-weight: normal; font-family: Calibri; font-size: 10pt" id=' + $table_id + '>';
    html += generate_table_caption(company, month, year, title);
    html += generate_table_head(month, year, mode);
    html += generate_table_body(dataset, curr_month_year, prev_month_year, mode);
    html += '</table>';
    html += '</div>';

    return html;
}

// shortcut that generates the consolidated table without extra adjustments
function generate_balance_sheet(company, month, year, dataset, curr_month_year, prev_month_year, $table_id, mode) {
    var html = generate_table_css();
    
    // the table containing all the data in html format
    html += '<div id="data">';
    html += '<table style="font-weight: normal; font-family: Calibri; font-size: 10pt" id=' + $table_id + '>';
    html += generate_table_caption(company, month, year, "Consolidated Balance Sheet");
    html += generate_table_head(month, year, mode);
    html += generate_table_body(dataset, curr_month_year, prev_month_year, mode);
    html += '</table>';
    html += '</div>';

    return html;
}

// generates the css for the html
function generate_table_css() {
    var table_css = "";

    table_css += '<style>';
    table_css += '    .table-data-right { font-family: Calibri; font-weight: normal; text-align: right; }';
    table_css += '    .table-data-left  { font-family: Calibri; font-weight: normal; text-align: left;  }';
    table_css += '</style>';

    return table_css;
}

// generates the table's caption on top
function generate_table_caption(company, month, year, title) {
    var table_caption = "";

    table_caption += '<caption style="text-align: left;">';
    table_caption += '    <span style="font-family: Calibri; font-size: 10pt; text-align: left;">' + company + '</br></span>';
    table_caption += '    <span style="font-family: Calibri; font-size: 10pt; text-align: left;">' + title + '</br></span>';
    table_caption += '    <span style="font-family: Calibri; font-size: 10pt; text-align: left;">' + month + '&nbsp;31,&nbsp;' + year + '</span>';
    table_caption += '</caption>';

    return table_caption;
}

// generates the table's column names
function generate_table_head(month, year, mode) {
    var html = "";
    var blank_column = indent + indent + "&nbsp;&nbsp;&nbsp;";

    if (mode == "income_statement") {
        html += '<thead>';
        html += '<tr style="border-top: 1px solid black; border-bottom: 1px solid black;">';
        html += '<th style="text-align: right; font-size: 10pt" colspan=3></th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + get_formatted_date(month, year, 0) + '</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + blank_column + '</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + get_formatted_date(month, year, 1) + '</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + blank_column + '</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + 'YTD ' + year.toString().slice(-2) + '</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + blank_column + '</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + 'YTD ' + (parseInt(year) - 1).toString().slice(-2) + '</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + blank_column + '</th>';
        html += '</tr>';
        html += '</thead>';
    } else if (mode == "trailing_12_months") {
        var ttm_period = get_ttm_period(month.slice(0, 3).toLowerCase() + "_" + year);

        html += '<thead>';
        html += '<tr style="border-top: 1px solid black; border-bottom: 1px solid black;">';
        html += '<th style="text-align: right; font-size: 10pt" colspan=3></th>';

        for(let i = 0; i < ttm_period.length; i++)
            html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + get_formatted_date(ttm_period[i], ttm_period[i], 0) + '</th>';

        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + indent + 'Total</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + indent + '%</th>';
        html += '</tr>';
        html += '</thead>';		
    } else if (mode == "balance_sheet") {
        html += '<thead>';
        html += '<tr style="border-top: 1px solid black; border-bottom: 1px solid black;">';
        html += '<th style="text-align: right; font-size: 10pt" colspan=3></th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + indent + get_formatted_date(month, year, 0) + '</th>';
        html += '<th style="text-align: right; font-size: 10pt" colspan=1>' + indent + indent + get_formatted_date(month, year, 1) + '</th>';
        html += '</tr>';
        html += '</thead>';
    }
    return html;
}

// generates the table's rows
function generate_table_body(dataset, curr_month_year, prev_month_year, mode) {
    var html = ""; // holds the html that is returned
    
    html += '<tbody>'; // start html table body
    if (mode == "income_statement" || mode == "trailing_12_months") {
        var categories = [
        // income
            "Product Sales", "Other Income",
        // cogs section
            "Cost of Goods Sold",
        // expenses
            "Operating Expenses", "Other Expenses",
        ];

        for (let i = 0; i < categories.length; i++) {
            if (i == 0)
                html += append_group_row("Income", true);
            if (i == 3) 
                html += append_group_row("Expenses", true);

            if (category_exists(dataset, categories[i]))
                html += get_category_rows(categories[i], dataset, curr_month_year, prev_month_year, mode) + "<tr></tr>";
            else 
                if (console_log) 
                    console.log("[!] " + categories[i] + " does not exists");
        }
    } else if (mode == "balance_sheet") {
        var categories = [
        // assets
            "Accounts Receivable",  "Bank", "Inventory", "Other Current Assets", "Fixed Assets",
        // liabilities
            "Accounts Payable", "Current Liabilities", "Duties and Taxes", "Long-Term Liabilities",
        ];

        for (let i = 0; i < categories.length; i++) {
            let found_category = category_exists(dataset, categories[i]);

            if (i == 0) {
                html += append_group_row("ASSETS", true) + append_group_row(indent + "Current Assets", false);
            } else if (i == 5) {

                for (let i = 0; i < dataset.length; i++) {
                    if (dataset[i]["account"] == "Assets") {
                        var total_assets = [
                            dataset[i][curr_month_year],
                            dataset[i][prev_month_year]
                        ];
                        html += append_total_row("Assets", total_assets, mode, true);

                        break;
                    }
                }

                html += "<tr></tr>";
                html += append_group_row("LIABILITIES", true);
            } else {
                if (found_category)
                    html += '<tr></tr>';
            }

            if (found_category) {
                html += get_category_rows(categories[i], dataset, curr_month_year, prev_month_year, mode);
            } else {
                if (console_log) 
                    console.log("[!] " + categories[i] + " does not exists");
            }
        }

        for (let i = 0; i < dataset.length; i++) {
            if (dataset[i]["account"] == "Liabilities") {
                var total_liabilities = [
                    dataset[i][curr_month_year],
                    dataset[i][prev_month_year]
                ];
                html += append_total_row("Liabilities", total_liabilities, mode, true);

                break;
            }
        }

    }
    html += '</tbody>'; // end html table body

    return html;
}


// ============================================================================================================================================
// GETTER FUNCTIONS
// ============================================================================================================================================


// formats the name of the account that's passed as arg // fixes capitalizations and removal of numbers
function get_formatted_name(account_object) {
    var account = "";

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

// converts dates into the format "MMM YY" // offset substracts the number from the year
function get_formatted_date(month, year, offset) {
    return ("&nbsp;" + month.toString().slice(0, 3).toUpperCase() + " " + (parseInt(year.slice(-2)) - offset).toString());
}

// formats the arg number such that negative numbers are surrounded by brackets // i.e., -1 to (1)
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

// arg must be in the format mmm_yyyy // returns an array containing the current mmm_yyyy and the preceding 11 months
function get_ttm_period(curr_month_year) {
    let curr_month = curr_month_year.slice(0, 3); // mmm
    let curr_year = curr_month_year.slice(4); // yyyy

    let month_names = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    let start_month = ""
    let start_year = parseInt(curr_year) - 1;

    let j = 0;
    while (month_names[j] != curr_month)
        j++;

    start_month = month_names[j++];

    let ttm_period = [];

    for (let i = j; i < (j+12); i++) {
        if (i > 1 && i % 12 == 0)
            start_year = parseInt(curr_year);

        ttm_period.push(month_names[i % 12] + "_" + start_year);
    }

    return ttm_period;
}

// compresses and returns the rows as print_groups[] that fall under the "catergory_name" arg // switch between modes using the "trailing_12_months" boolean
function get_category_rows(category_name, dataset, curr_month_year, prev_month_year, mode) {

    let index = 0;
    let html = "";
    let account = "";
    let print_groups = [];
    let ttm_period = get_ttm_period(curr_month_year);
    let category_total = get_category_total(category_name, dataset, curr_month_year, prev_month_year, mode);

    // find the beginning of this category and keep the index
    while (dataset[index]["account"] != category_name && index < dataset.length) 
        index++;

    // append the group header before the rest of the rows
    html += append_group_row(get_formatted_name(dataset[index]), false);

    // we need to move to the next index because the curernt index is the header itself
    index++;

    if (mode == "income_statement") {
        // gather each row's data under the current category
        // also compresses the accounts based on print groups
        // if a print group already exists in print_groups[], sum up their values 
        // if it does not exist already, append that group to print_groups[] along with its data
        while (dataset[index]["parent_account"] == category_name) {
            account = dataset[index]["account"];

            // this section compares the current print group against existing print groups
            let group_found = false;
            for (let j = 0; j < print_groups.length; j++) {
                if (print_groups[j].account == account) {
                    print_groups[j].curr_data += dataset[index][curr_month_year];
                    print_groups[j].prev_data += dataset[index][prev_month_year];
                    print_groups[j].curr_ytd  += dataset[index]["total"];
                    print_groups[j].prev_ytd  += dataset[index]["prev_year_total"];
                    print_groups[j].indent    =  dataset[index]["indent"];
                    print_groups[j].is_group  =  dataset[index]["is_group"];

                    group_found = true;
                    break;
                }
            }

            // if the print group was not found, append it to the end
            if (!group_found) {
                // create an object containing the info and push() to print_groups[]
                print_groups.push({
                    "account"   : account,
                    "curr_data" : dataset[index][curr_month_year],
                    "prev_data" : dataset[index][prev_month_year],
                    "curr_ytd"  : dataset[index]["total"],
                    "prev_ytd"  : dataset[index]["prev_year_total"],
                    "indent"    : dataset[index]["indent"],
                    "is_group"  : dataset[index]["is_group"]
                });
            }
            
            index++;
            
            // break the loop if no more rows exist in the source array
            if (!dataset[index]) 
                break;
        }
        
        var data = [];
        // adds each row's gathered data to the html
        for (let i = 0; i < print_groups.length; i++) {
            account = get_formatted_name(print_groups[i]);

            data = [
                print_groups[i].curr_data,
                print_groups[i].prev_data,
                print_groups[i].curr_ytd,
                print_groups[i].prev_ytd
            ];

            html += append_data_row(category_total, account, data, mode);
        }
    
        // appends a row containing the total values for the current category
        html += append_total_row(category_name, category_total, mode, false);

    } else if (mode == "trailing_12_months") {
        while (dataset[index]["parent_account"] == category_name) {
            account = dataset[index]["account"];

            // this section compares the current print group against existing print groups
            let group_found = false;
            for (let j = 0; j < print_groups.length; j++) {
                if (print_groups[j].account == account) {
                    print_groups[j]["ttm_00"]   += dataset[index][ttm_period[0]],
                    print_groups[j]["ttm_01"]   += dataset[index][ttm_period[1]],
                    print_groups[j]["ttm_02"]   += dataset[index][ttm_period[2]],
                    print_groups[j]["ttm_03"]   += dataset[index][ttm_period[3]],
                    print_groups[j]["ttm_04"]   += dataset[index][ttm_period[4]],
                    print_groups[j]["ttm_05"]   += dataset[index][ttm_period[5]],
                    print_groups[j]["ttm_06"]   += dataset[index][ttm_period[6]],
                    print_groups[j]["ttm_07"]   += dataset[index][ttm_period[7]],
                    print_groups[j]["ttm_08"]   += dataset[index][ttm_period[8]],
                    print_groups[j]["ttm_09"]   += dataset[index][ttm_period[9]],
                    print_groups[j]["ttm_10"]   += dataset[index][ttm_period[10]],
                    print_groups[j]["ttm_11"]   += dataset[index][ttm_period[11]],
                    print_groups[j]["total"]    += dataset[index]["total"];
                    print_groups[j]["indent"]   =  dataset[index]["indent"];
                    print_groups[j]["is_group"] =  dataset[index]["is_group"];

                    group_found = true;
                    break;
                }
            }

            // if the print group was not found, append it to the end
            if (!group_found) {
                // create an object containing the info and push() to print_groups[]
                print_groups.push({
                    "account"  : account,
                    "ttm_00"   : dataset[index][ttm_period[0]],
                    "ttm_01"   : dataset[index][ttm_period[1]],
                    "ttm_02"   : dataset[index][ttm_period[2]],
                    "ttm_03"   : dataset[index][ttm_period[3]],
                    "ttm_04"   : dataset[index][ttm_period[4]],
                    "ttm_05"   : dataset[index][ttm_period[5]],
                    "ttm_06"   : dataset[index][ttm_period[6]],
                    "ttm_07"   : dataset[index][ttm_period[7]],
                    "ttm_08"   : dataset[index][ttm_period[8]],
                    "ttm_09"   : dataset[index][ttm_period[9]],
                    "ttm_10"   : dataset[index][ttm_period[10]],
                    "ttm_11"   : dataset[index][ttm_period[11]],
                    "total"    : dataset[index]["total"],
                    "indent"   : dataset[index]["indent"],
                    "is_group" : dataset[index]["is_group"]
                });
            }
            
            index++;
            
            // break the loop if no more rows exist in the source array
            if (!dataset[index]) 
                break;
        }
        
        var data = [];
        // adds each row's gathered data to the html
        for (let i = 0; i < print_groups.length; i++) {
            account = get_formatted_name(print_groups[i]);

            data = [];
            data.push(print_groups[i]["ttm_00"]);
            data.push(print_groups[i]["ttm_01"]);
            data.push(print_groups[i]["ttm_02"]);
            data.push(print_groups[i]["ttm_03"]);
            data.push(print_groups[i]["ttm_04"]);
            data.push(print_groups[i]["ttm_05"]);
            data.push(print_groups[i]["ttm_06"]);
            data.push(print_groups[i]["ttm_07"]);
            data.push(print_groups[i]["ttm_08"]);
            data.push(print_groups[i]["ttm_09"]);
            data.push(print_groups[i]["ttm_10"]);
            data.push(print_groups[i]["ttm_11"]);
            data.push(print_groups[i]["total"]);

            html += append_data_row(category_total, account, data, mode);
        }

        // appends a row containing the total values for the current category
        html += append_total_row(category_name, category_total, mode, false);
    } else if (mode == "balance_sheet") {
        while (dataset[index]["parent_account"] == category_name) {
            account = dataset[index]["account"];

            // this section compares the current print group against existing print groups
            let group_found = false;
            for (let j = 0; j < print_groups.length; j++) {
                if (print_groups[j].account == account) {
                    print_groups[j].curr_data += dataset[index][curr_month_year];
                    print_groups[j].prev_data += dataset[index][prev_month_year];
                    print_groups[j].indent    =  dataset[index]["indent"];
                    print_groups[j].is_group  =  dataset[index]["is_group"];

                    group_found = true;
                    break;
                }
            }

            // if the print group was not found, append it to the end
            if (!group_found) {
                // create an object containing the info and push() to print_groups[]
                print_groups.push({
                    "account"   : account,
                    "curr_data" : dataset[index][curr_month_year],
                    "prev_data" : dataset[index][prev_month_year],
                    "indent"    : dataset[index]["indent"],
                    "is_group"  : dataset[index]["is_group"]
                });
            }
            
            index++;
            
            // break the loop if no more rows exist in the source array
            if (!dataset[index]) 
                break;
        }
        
        var data = [];
        // adds each row's gathered data to the html
        for (let i = 0; i < print_groups.length; i++) {
            account = get_formatted_name(print_groups[i]);

            data = [
                print_groups[i].curr_data,
                print_groups[i].prev_data,
            ];

            html += append_data_row(category_total, account, data, mode);
        }
    
        // appends a row containing the total values for the current category
        html += append_total_row(category_name, category_total, mode, false);
    }

    if (console_log) 
        console.log(" --> Appended " + category_name);

    return html;
}

// finds the total amount per year based on the category name passed // switch between modes using the "trailing_12_months" boolean
function get_category_total(category_name, dataset, curr_month_year, prev_month_year, mode) {
    if (console_log) 
        console.log("\t[" + category_name + "] calculating total");

    let nf = new Intl.NumberFormat('en-US');
    var ttm_period = get_ttm_period(curr_month_year);
    // values for:     [jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, tot]
    var total_values = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]; // array of totals
    var index = 0;
    
    // find the beginning of this category and keep the index
    while (dataset[index]["account"] != category_name && index < dataset.length) 
        index++;
    index++;

    var current_indent = dataset[index]["indent"];
    
    if (mode == "income_statement") {
        // everything under this subgroup is summed together into the array
        while (dataset[index]["indent"] == current_indent && index < dataset.length) {
            total_values[0] += dataset[index][curr_month_year];
            total_values[1] += dataset[index][prev_month_year];
            total_values[2] += dataset[index]["total"];
            total_values[3] += dataset[index]["prev_year_total"];

            index++;

            // break if end of array
            if (!dataset[index])
                break;
        }
    } else if (mode == "trailing_12_months") {
        while (dataset[index]["indent"] == current_indent && index < dataset.length) {
            for (let k = 0; k < total_values.length - 1; k++)
                total_values[k] += dataset[index][ttm_period[k]];

            total_values[12] += dataset[index]["total"];

            index++;

            // break if end of array
            if (!dataset[index])
                break;
        }
    } else if (mode == "balance_sheet") {
        while (dataset[index]["indent"] == current_indent && index < dataset.length) {
            
            total_values[0] += dataset[index][curr_month_year];
            total_values[1] += dataset[index][prev_month_year];

            index++;

            // break if end of array
            if (!dataset[index])
                break;
        }
    }
    
    // round down all the values before returning the array
    for (let j = 0; j < total_values.length; j++)
        nf.format(Math.floor(total_values[j]));
    
    if (console_log) 
        console.log("\t[" + category_name + "] total calculated");

    return total_values;
}

// send the entire array as "dataset" // checks if the category passed as "category_name" exists in the dataset
function category_exists(dataset, category_name) {
    var category_exists = false;

    for (let i = 0; i < dataset.length; i++) {
        if (dataset[i]["account"].slice(0, category_name.length) == category_name){
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


// appends the name that is passed as argument // no data is appended -- used for headers and such
function append_group_row(account, is_root) {
    var html = "";

    html += '<tr>';
    if (is_root)
        html += '<td class="table-data-right" style="font-weight: bold; font-size: 10pt" colspan=3>' + account.toUpperCase() + '</td>';
    else
        html += '<td class="table-data-right" style="font-size: 10pt" colspan=3>' + account + '</td>';
    html += '</tr>';

    return html
}

// appends each row's data (used inside get_category_rows()), along with the percentage // switch between modes using the "trailing_12_months" boolean
function append_data_row(total_array, account, data, mode) {
    let nf = new Intl.NumberFormat('en-US');
    var html = "";
    var values = [];
    var percentages = [];
        
    html += '<tr>';
    html += '<td class="table-data-right" style="font-size: 10pt" colspan=3>' + account + '</td>';

    if (mode == "income_statement") {
        for (let i = 0; i < data.length; i++) {
            // round down and format the number to 2 decimal places
            values.push((nf.format(Math.floor(data[i]))));
            // get_formatted_number() replaces minus symbols with brackets when it the global flag is true
            percentages.push(get_formatted_number(((data[i] * 100) / total_array[i]).toFixed(2)) + "%");
        }

        for (let i = 0; i < 4; i++) {
            html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>' + get_formatted_number(values[i]) + '</td>';

            if (percentages[i].toString().slice(0, -1) == "NaN" || percentages[i].toString().slice(0, -1) == "100.00") {
                html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>100%</td>';
            } else if (percentages[i].toString().slice(0, -1) == "0.00") {
                html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>0%</td>';
            } else {
                html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>' + percentages[i] + '</td>';
            }
        }
    } else if (mode == "trailing_12_months") {
        for (let i = 0; i < data.length; i++)			
            values.push((nf.format(Math.floor(data[i])))); // round down and format the number to 2 decimal places

        // get_formatted_number() replaces minus symbols with brackets when it the global flag is true
        let percentage = (get_formatted_number(((data[data.length - 1] * 100) / total_array[total_array.length - 1]).toFixed(2)) + "%");

        for (let i = 0; i < data.length; i++)
            html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>' + get_formatted_number(values[i]) + '</td>';

        if (percentage.toString().slice(0, -1) == "NaN" || percentage.toString().slice(0, -1) == "100.00") {
            html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>100%</td>';
        } else if (percentage.toString().slice(0, -1) == "0.00") {
            html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>0%</td>';
        } else {
            html += '<td class="table-data-right" style="text-align: right; font-size: 10pt" colspan=1>' + percentage + '</td>';
        }
    } else if (mode == "balance_sheet") {
        html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + (nf.format(Math.floor(data[0]))) + '</td>';
        html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + (nf.format(Math.floor(data[1]))) + '</td>';
    }

    html += '</tr>';
        
    return html;
}

// appends the category's data under each year (used inside get_category_rows()), along with the percentage // switch between modes using the "trailing_12_months" boolean
function append_total_row(category_name, category_total, mode, is_root) { 
    let nf = new Intl.NumberFormat('en-US');
    
    var html = "";
    var values = [];

    for(let i = 0; i < category_total.length; i++)
        values.push(nf.format(Math.floor(category_total[i])));

    html += '<tr style="border-top: 1px solid black">';

    if (is_root)
        html += '<td class="table-data-right" style="font-weight: bold; font-size: 10pt" colspan=3>' + 'TOTAL ' + category_name.toUpperCase() + '</td>';
    else
        html += '<td class="table-data-right" style="font-size: 10pt" colspan=3>' + indent + 'Total ' + category_name + '</td>';


    if (mode == "income_statement") {
        for (let i = 0; i < 4; i++) {
            html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + values[i] + '</td>';
            html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>100%</td>';
        }
    } else if (mode == "trailing_12_months") {
        for (let i = 0; i < category_total.length; i++)
            html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + values[i] + '</td>';

        html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>100%</td>';
    } else if (mode == "balance_sheet") {
        for (let i = 0; i < 2; i++)
            html += '<td class="table-data-right" style="font-size: 10pt" colspan=1>' + values[i] + '</td>';
    }

    html += '</tr>';

    return html;
}


// ============================================================================================================================================
// EXPORT FUNCTION
// ============================================================================================================================================


// assign names to each sheet based on cost center // exports the excel file
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

            switch (center_number) {
                case '0_IS': SheetName = "IS - Consolidated"; break;
                case '1_IS': SheetName = "IS - White-Wood Corporate"; break;
                case '2_IS': SheetName = "IS - White-Wood Distributors"; break;
                case '3_IS': SheetName = "IS - Forest Products"; break;
                case '6_IS': SheetName = "IS - Endeavours"; break;

                case '0_TTM': SheetName = "TTM - Consolidated"; break;
                case '1_TTM': SheetName = "TTM - White-Wood Corporate"; break;
                case '2_TTM': SheetName = "TTM - White-Wood Distributors"; break;
                case '3_TTM': SheetName = "TTM - Forest Products"; break;
                case '6_TTM': SheetName = "TTM - Endeavours"; break;
                
                case '0_BS': SheetName = "BS - Consolidated"; break;

                default: SheetName = "Sheet"; break;
            }
            
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
