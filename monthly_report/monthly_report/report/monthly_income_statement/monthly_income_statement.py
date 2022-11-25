# Copyright (c) 2022, Farabi Hussain

import frappe, calendar, copy, datetime, json
from frappe import _, qb, scrub
from frappe.query_builder import CustomFunction
from frappe.query_builder.functions import Max
from frappe.utils import date_diff, flt, getdate, cint, flt
from erpnext.controllers.queries import get_match_cond
from erpnext.accounts.report.financial_statements import *
from erpnext.stock.utils import get_incoming_rate
from collections import OrderedDict


global_fiscal_year = 0
curr_year = ""
prev_year = ""

## ======================================================================
## GATHER + STRUCTURE DATA WITH SQL
## ======================================================================

## main function for the custom app
def execute(filters = None):
    make_chart = False

    if make_chart:
        period  = get_period(filters.to_fiscal_year, filters.periodicity, filters.period_end_month, company = filters.company)
        income  = get_data(filters.period_end_month, filters.to_fiscal_year, filters.company, "Income", "Credit", period, filters = filters, accumulated_values = filters.accumulated_values, ignore_closing_entries = True, ignore_accumulated_values_for_fy = True)
        expense = get_data(filters.period_end_month, filters.to_fiscal_year, filters.company, "Expense", "Debit", period, filters = filters, accumulated_values = filters.accumulated_values, ignore_closing_entries = True, ignore_accumulated_values_for_fy = True)
        columns = get_columns(filters.period_end_month, filters.to_fiscal_year, filters.periodicity, period)

        data = []
        data.extend(income or [])
        data.extend(expense or [])

        return columns, data, None, None, None
    else:
        return None, None, None, None, None



## overriden from financial_statements.py -- goes through the columns and appends them based on the user-selected month
def get_columns(period_end_month, period_end_year, periodicity, period_list):
    columns = [{"fieldname": "account", "label": _("Account"), "fieldtype": "Link", "options": "Account", "width": 325}]
    end_month_and_year = (period_end_month[0:3] + " " + period_end_year)

    fiscal_year_start_stamp = ((int(period_end_year) - 1) * 100) + int(list(calendar.month_abbr).index(period_end_month[0:3]))

    for period in reversed(period_list):
        current_timestamp = ((int(period.label[4:8])) * 100) + int(list(calendar.month_abbr).index(period.label[0:3])) # current index in timestamp

        if (current_timestamp >= fiscal_year_start_stamp): # check the year
            if (period.label[0:3] == end_month_and_year[0:3]): # check the month
                columns.append({"fieldname": period.key, "label": period.label, "fieldtype": "Currency", "options": "currency", "width": 175})

    columns.append({"fieldname": "total", "label": _("YTD " + str(period_end_year)), "fieldtype": "Currency", "width": 175})
    columns.append({"fieldname": "prev_year_total", "label": _("YTD " + str(int(period_end_year)-1)), "fieldtype": "Currency", "width": 175})

    return columns



## overriden from financial_statements.py -- returns the timeframe for which this report is generated
def get_period(to_fiscal_year, periodicity, period_end_month, accumulated_values=False, company=None, reset_period_on_fy_change=True, ignore_fiscal_year=False):
    # Get a list of dict {"from_date": from_date, "to_date": to_date, "key": key, "label": label}
    # Periodicity can be (Yearly, Quarterly, Monthly)

    from_fiscal_year = str(int(to_fiscal_year) - 1)# we want to run this on the same fiscal year, so from_fiscal_year = to_fiscal_year

    # by default it gets the start month of the fiscal year, which can be different from January
    # but the first column should be the same column as the selected month, which may be from before the current fiscal year
    # to circumvent this, we pick months from the beginning of the calendar year if its before the fiscal year
    fiscal_year = get_fiscal_year_data(from_fiscal_year, to_fiscal_year)
    global global_fiscal_year
    global_fiscal_year = fiscal_year

    validate_fiscal_year(fiscal_year, from_fiscal_year, to_fiscal_year)
    selected_month_in_int = list(calendar.month_abbr).index(period_end_month[0:3]) # convert the selected month into int
    fiscal_starting_month_in_int = int(fiscal_year.year_start_date.strftime("%m")) # convert the fiscal year's starting month into int

    if (selected_month_in_int < fiscal_starting_month_in_int):
        build_start_year_and_date = (str(int(from_fiscal_year)-1) + '-' + str(selected_month_in_int) + '-01')
        year_start_date = getdate(build_start_year_and_date)
    else:
        year_start_date = getdate(fiscal_year.year_start_date)

    fiscal_year_start_month = fiscal_year.year_start_date
    year_end_date = getdate(fiscal_year.year_end_date)

    months_to_add = {"Yearly": 12, "Half-Yearly": 6, "Quarterly": 3, "Monthly": 1}[periodicity]
    period_list = []

    start_date = year_start_date
    months = get_months(year_start_date, year_end_date)

    ######################################################################################################
    for i in range(months):
        period = frappe._dict({"from_date": start_date})
        to_date = add_months(start_date, months_to_add)
        start_date = to_date
        to_date = add_days(to_date, -1) # Subtract one day from to_date, as it may be first day in next fiscal year or month

        if (to_date <= year_end_date): 
            period.to_date = to_date # the normal case
        else:
            period.to_date = year_end_date # if a fiscal year ends before a 12 month period

        if not ignore_fiscal_year:
            period.to_date_fiscal_year = get_fiscal_year(period.to_date, company=company)[0]
            period.from_date_fiscal_year_start_date = get_fiscal_year(period.from_date, company=company)[1]

        period_list.append(period)

        if period.to_date == year_end_date:
            break

    ######################################################################################################
    # common processing
    
    for opts in period_list:
        key = opts["to_date"].strftime("%b_%Y").lower()
        if periodicity == "Monthly" and not accumulated_values:
            label = formatdate(opts["to_date"], "MMM YYYY")
        else:
            if not accumulated_values:
                label = get_label(periodicity, opts["from_date"], opts["to_date"])
            else:
                if reset_period_on_fy_change:
                    label = get_label(periodicity, opts.from_date_fiscal_year_start_date, opts["to_date"])
                else:
                    label = get_label(periodicity, period_list[0].from_date, opts["to_date"])

        opts.update({
            "key": key.replace(" ", "_").replace("-", "_"),
            "label": label,
            "year_start_date": year_start_date,
            "year_end_date": year_end_date,
        })

    return period_list



## overriden from financial_statements.py
def get_data(period_end_month, period_end_year, company, root_type, balance_must_be, period_list, filters=None, accumulated_values=1, only_current_fiscal_year=True, ignore_closing_entries=False, ignore_accumulated_values_for_fy=False, total=True):
    end_month_and_year = (period_end_month[0:3] + " " + period_end_year)
    accounts = get_accounts(company, root_type)
    
    if not accounts:
        return None

    accounts, accounts_by_name, parent_children_map = filter_accounts(accounts)
    company_currency = get_appropriate_currency(company, filters)
    gl_entries_by_account = {}

    # extracts the root of the trees "Income" and "Expenses"
    # only two elements in this dict
    print("\tgetting list of accounts")
    accounts_list = frappe.db.sql(
        """
        SELECT 
            lft,
            rgt 
        FROM 
            tabAccount
        WHERE 
            root_type=%s 
            AND IFNULL(parent_account, '') = ''
        """,
        root_type,
        as_dict = True
    )

    # for both of the trees, extract the leaves and populate gl_entries_by_account
    for root in accounts_list:
        set_gl_entries_by_account(company, period_list[0]["year_start_date"] if only_current_fiscal_year else None, period_list[-1]["to_date"], root.lft, root.rgt, filters, gl_entries_by_account, ignore_closing_entries=ignore_closing_entries)

    calculate_values(accounts_by_name, gl_entries_by_account, period_list, accumulated_values, ignore_accumulated_values_for_fy) ## function imported from financial_statements.py
    accumulate_values_into_parents(accounts, accounts_by_name, period_list)                                                      ## function imported from financial_statements.py
    out = prepare_data(end_month_and_year, accounts, balance_must_be, period_list, company_currency)
    out = filter_out_zero_value_rows(out, parent_children_map)

    for data in out:
        if data: 
            if data.account[-5:] == " - WW":
                data.account = (data.account)[:-5]
            if data.parent_account[-5:] == " - WW":
                data.parent_account = (data.parent_account)[:-5]

    return out



## overriden from financial_statements.py -- calculates the dollar values to be put in each cell, one row at a time 
def prepare_data(end_month_and_year, accounts, balance_must_be, period_list, company_currency):

    data = []
    year_start_date = period_list[0]["year_start_date"].strftime("%Y-%m-%d")
    year_end_date = period_list[-1]["year_end_date"].strftime("%Y-%m-%d")

    # variables for current fiscal year calculation
    global global_fiscal_year
    fiscal_year_start_month_in_int = int(global_fiscal_year.year_start_date.strftime("%m"))
    fiscal_year_in_int = int(global_fiscal_year.year_start_date.strftime("%Y"))
    fiscal_year_stamp = ((fiscal_year_in_int + 1) * 100) + fiscal_year_start_month_in_int

    # variables for previous fiscal year calculation
    prev_fiscal_year_end_month = list(calendar.month_abbr).index(end_month_and_year[0:3])
    prev_fiscal_year_start = ((fiscal_year_in_int) * 100) + fiscal_year_start_month_in_int
    prev_fiscal_year_end = ((fiscal_year_in_int + 1) * 100) + prev_fiscal_year_end_month

    counter = len(accounts)
    current = 0
    for account in accounts:
        has_value = False
        total = 0
        prev_year_total = 0
        print_group = frappe.db.sql(
            """
            SELECT  print_group 
            FROM    tabAccount 
            WHERE   name = %s
            """,
            account.name
        )

        row = frappe._dict(
            {
                "account": _(account.name),
                "parent_account": _(account.parent_account) if account.parent_account else "",
                "indent": flt(account.indent),
                "year_start_date": year_start_date,
                "year_end_date": year_end_date,
                "currency": company_currency,
                "include_in_gross": account.include_in_gross,
                "account_type": account.account_type,
                "is_group": account.is_group,
                "opening_balance": account.get("opening_balance", 0.0) * (1 if balance_must_be == "Debit" else -1),
                "account_name": (
                    "%s - %s" % (_(account.account_number), _(account.account_name))
                    if account.account_number
                    else _(account.account_name)
                ),
            }
        )

        for period in period_list:
            if account.get(period.key) and balance_must_be == "Credit":
                account[period.key] *= -1 # change sign based on Debit or Credit, since calculation is done using (debit - credit)

            row[period.key] = flt(account.get(period.key, 0.0), 3)

            # ignore zero values
            if abs(row[period.key]) >= 0.005:
                has_value = True

                current_month_in_int = list(calendar.month_abbr).index(period.label[0:3]) # convert month name to month number
                current_year_in_int = int(period.label[4:8])                              # period.label contains the date and time
                current_year_stamp = (current_year_in_int * 100) + current_month_in_int   # creates a timestamp in the format yyyymm for date comparison

                if (current_year_stamp >= fiscal_year_stamp):
                    total += flt(row[period.key])

                if (prev_fiscal_year_start <= current_year_stamp and current_year_stamp <= prev_fiscal_year_end):
                    prev_year_total += flt(row[period.key])
            
            if (period.label == end_month_and_year):
                break

        if (row["is_group"] == False): 
            row["account"] = print_group[0][0]

        if (row["account"] == ""):
            row["account"] = row["account_name"]

        row["has_value"] = has_value
        row["total"] = total
        row["print_group"] = print_group[0][0]
        row["prev_year_total"] = prev_year_total

        data.append(row)
        current += 1
        if ((current/counter * 100) % 5 < 0.25): print("\tpreparing data " + str(int(current/counter * 100)) + "%")

    return data



## overriden from financial_statements.py -- 
def set_gl_entries_by_account(company, from_date, to_date, root_lft, root_rgt, filters, gl_entries_by_account, ignore_closing_entries=False):
    # Returns a dict like { "account": [gl entries], ... }
    additional_conditions = get_additional_conditions(from_date, ignore_closing_entries, filters)

    accounts = frappe.db.sql_list(
        """
        SELECT
            name
        FROM 
            `tabAccount`
        WHERE 
            lft >= %s
            AND rgt <= %s
            AND company = %s
        """,
        (root_lft, root_rgt, company)
    )

    if accounts:
        additional_conditions += " AND account IN ({})".format(
            ", ".join(frappe.db.escape(account) for account in accounts)
        )

        gl_filters = {
            "company": company,
            "from_date": from_date,
            "to_date": to_date,
            "finance_book": cstr(filters.get("finance_book")),
        }

        if filters.get("include_default_book_entries"):
            gl_filters["company_fb"] = frappe.db.get_value("Company", company, "default_finance_book")

        for key, value in filters.items():
            if value:
                gl_filters.update({key: value})

        distributed_cost_center_query = ""

        if filters and filters.get("cost_center"):
            distributed_cost_center_query = (
                """
                UNION ALL
                SELECT 
                    posting_date,
                    account,
                    debit*(DCC_allocation.percentage_allocation/100) AS debit,
                    credit*(DCC_allocation.percentage_allocation/100) AS credit,
                    is_opening,
                    fiscal_year,
                    debit_in_account_currency*(DCC_allocation.percentage_allocation/100) AS debit_in_account_currency,
                    credit_in_account_currency*(DCC_allocation.percentage_allocation/100) AS credit_in_account_currency,
                    account_currency
                FROM 
                    `tabGL Entry`,
                    (
                        SELECT 
                            parent, 
                            sum(percentage_allocation) AS percentage_allocation
                        FROM 
                            `tabDistributed Cost Center`
                        WHERE 
                            cost_center IN %(cost_center)s
                            AND parent NOT IN %(cost_center)s
                        GROUP BY 
                            parent
                    ) AS DCC_allocation
                WHERE 
                    company=%(company)s
                    {additional_conditions}
                    AND posting_date <= %(to_date)s
                    AND is_cancelled = 0
                    AND cost_center = DCC_allocation.parent
                """.format(
                    additional_conditions = additional_conditions.replace("AND cost_center IN %(cost_center)s ", "")
                )
            )

        gl_entries = frappe.db.sql(
            """
            SELECT 
                posting_date,
                account,
                debit,
                credit,
                is_opening,
                fiscal_year,
                debit_in_account_currency,
                credit_in_account_currency,
                account_currency 
            FROM 
                `tabGL Entry`
            WHERE 
                company=%(company)s
                {additional_conditions}
                AND posting_date <= %(to_date)s
                AND is_cancelled = 0
                {distributed_cost_center_query}
            """.format(
                additional_conditions=additional_conditions,
                distributed_cost_center_query=distributed_cost_center_query,
            ),
            gl_filters,
            as_dict = True,
        )

        if filters and filters.get("presentation_currency"):
            convert_to_presentation_currency(gl_entries, get_currency(filters), filters.get("company"))

        for entry in gl_entries:
            gl_entries_by_account.setdefault(entry.account, []).append(entry)



## ======================================================================
## FUNCTION WORKING WITH JAVASCRIPT
## ======================================================================

@frappe.whitelist()
def get_records(filters):
    filterDt = json.loads(filters)
    filters = frappe._dict(filterDt or {})

    if not filters:
        return [], [], None, []

    print("(0) validating filters")
    validate_filters(filters)
    print("(1) getting period")
    period  = get_period(filters.to_fiscal_year, filters.periodicity, filters.period_end_month, company = filters.company)
    print("(2) getting income")
    income  = get_data(filters.period_end_month, filters.to_fiscal_year, filters.company, "Income", "Credit", period, filters = filters, accumulated_values = filters.accumulated_values, ignore_closing_entries = True, ignore_accumulated_values_for_fy = True)
    print("(3) getting expense")
    expense = get_data(filters.period_end_month, filters.to_fiscal_year, filters.company, "Expense", "Debit", period, filters = filters, accumulated_values = filters.accumulated_values, ignore_closing_entries = True, ignore_accumulated_values_for_fy = True)
    print("(4) getting columns")
    columns = get_columns(filters.period_end_month, filters.to_fiscal_year, filters.periodicity, period)
    print("(5) download started")

    data = []
    data.extend(income or [])
    data.extend(expense or [])

    if not data:
        return [], [], None, []
    
    return columns, data



## validates the filters for the get_records() function
def validate_filters(filters):
    if not filters.period_end_month:
        frappe.throw(_("Please select a month."))
    if not filters.to_fiscal_year:
        frappe.throw(_("Please select a year."))
    if not filters.cost_center:
        frappe.throw(_("Please select at least one cost center."))
