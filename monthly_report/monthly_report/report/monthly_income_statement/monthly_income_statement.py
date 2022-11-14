# Copyright (c) 2022, Farabi Hussain

import frappe
import calendar
from frappe import _, scrub
from erpnext.accounts.report.financial_statements import *


global_fiscal_year = 0

######################################################################################################
## main function for the custom app
######################################################################################################
def execute(filters = None):
    curr_year_period  = get_period_list(filters.to_fiscal_year, filters.periodicity, filters.period_end_month, company = filters.company)
    curr_year_income  = get_data(filters.period_end_month, filters.to_fiscal_year, filters.company, "Income", "Credit", curr_year_period, filters = filters, accumulated_values = filters.accumulated_values, ignore_closing_entries = True, ignore_accumulated_values_for_fy = True)
    curr_year_expense = get_data(filters.period_end_month, filters.to_fiscal_year, filters.company, "Expense", "Debit", curr_year_period, filters = filters, accumulated_values = filters.accumulated_values, ignore_closing_entries = True, ignore_accumulated_values_for_fy = True)
    curr_year_columns = get_columns(filters.period_end_month, filters.to_fiscal_year, filters.periodicity, curr_year_period)

    data = []
    data.extend(curr_year_income or [])
    data.extend(curr_year_expense or [])
    columns = curr_year_columns

    return columns, data, None, None, None


######################################################################################################
## overriden from from financial_statements.py 
## -- goes through the columns and appends them based on the user-selected month
######################################################################################################
def get_columns(period_end_month, period_end_year, periodicity, period_list):
    columns = [{"fieldname": "account", "label": _("Account"), "fieldtype": "Link", "options": "Account", "width": 350}]
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


######################################################################################################
## overriden from from financial_statements.py
## -- returns the timeframe for which this report is generated
######################################################################################################
def get_period_list(to_fiscal_year, periodicity, period_end_month, accumulated_values=False, company=None, reset_period_on_fy_change=True, ignore_fiscal_year=False):
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


######################################################################################################
## overriden from from financial_statements.py
##
######################################################################################################
def get_data(period_end_month, period_end_year, company, root_type, balance_must_be, period_list, filters=None, accumulated_values=1, only_current_fiscal_year=True, ignore_closing_entries=False, ignore_accumulated_values_for_fy=False, total=True):
    end_month_and_year = (period_end_month[0:3] + " " + period_end_year)
    accounts = get_accounts(company, root_type)
    
    if not accounts:
        return None

    accounts, accounts_by_name, parent_children_map = filter_accounts(accounts)
    company_currency = get_appropriate_currency(company, filters)
    gl_entries_by_account = {}

    accounts_list = frappe.db.sql(
        """
        SELECT 
            lft,
            rgt,
            print_group 
        FROM 
            tabAccount
        WHERE 
            root_type=%s 
            AND IFNULL(parent_account, '') = ''
        """,
        root_type,
        as_dict=1,
    )

    for root in accounts_list:
        set_gl_entries_by_account(
            company,
            period_list[0]["year_start_date"] if only_current_fiscal_year else None,
            period_list[-1]["to_date"],
            root.lft,
            root.rgt,
            filters,
            gl_entries_by_account,
            ignore_closing_entries=ignore_closing_entries,
        )

    ## function imported from financial_statements.py   
    calculate_values(accounts_by_name, gl_entries_by_account, period_list, accumulated_values, ignore_accumulated_values_for_fy)
    ## function imported from financial_statements.py
    accumulate_values_into_parents(accounts, accounts_by_name, period_list)
    out = prepare_data(end_month_and_year, accounts, balance_must_be, period_list, company_currency)

    if out and total:
        add_total_row(end_month_and_year, out, root_type, balance_must_be, period_list, company_currency)

    out = filter_out_zero_value_rows(out, parent_children_map)

    return out


######################################################################################################
## overriden from financial_statements.py
## -- calculates the dollar values to be put in each cell 
######################################################################################################
def prepare_data(end_month_and_year, accounts, balance_must_be, period_list, company_currency):

    data = []
    year_start_date = period_list[0]["year_start_date"].strftime("%Y-%m-%d")
    year_end_date = period_list[-1]["year_end_date"].strftime("%Y-%m-%d")

    global global_fiscal_year
    fiscal_year_start_month_in_int = int(global_fiscal_year.year_start_date.strftime("%m"))
    fiscal_year_in_int = int(global_fiscal_year.year_start_date.strftime("%Y"))
    fiscal_year_stamp = ((fiscal_year_in_int + 1) * 100) + fiscal_year_start_month_in_int

    prev_fiscal_year_end_month = list(calendar.month_abbr).index(end_month_and_year[0:3])
    prev_fiscal_year_start = ((fiscal_year_in_int) * 100) + fiscal_year_start_month_in_int
    prev_fiscal_year_end = ((fiscal_year_in_int + 1) * 100) + prev_fiscal_year_end_month

    for d in accounts:
        # add to output
        has_value = False
        total = 0
        prev_year_total = 0
        print_group = ""
        row = frappe._dict(
            {
                "account": _(d.name),
                "parent_account": _(d.parent_account) if d.parent_account else "",
                "indent": flt(d.indent),
                "year_start_date": year_start_date,
                "year_end_date": year_end_date,
                "currency": company_currency,
                "include_in_gross": d.include_in_gross,
                "account_type": d.account_type,
                "is_group": d.is_group,
                "opening_balance": d.get("opening_balance", 0.0) * (1 if balance_must_be == "Debit" else -1),
                "account_name": (
                    "%s - %s" % (_(d.account_number), _(d.account_name))
                    if d.account_number
                    else _(d.account_name)
                ),
            }
        )

        for period in period_list:
            if d.get(period.key) and balance_must_be == "Credit":
                d[period.key] *= -1 # change sign based on Debit or Credit, since calculation is done using (debit - credit)

            row[period.key] = flt(d.get(period.key, 0.0), 3)

            # ignore zero values
            if abs(row[period.key]) >= 0.000:
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

        row["has_value"] = has_value
        row["total"] = total
    
        # remove value from the "Income" and "Expenses" rows
        if (d.name[:3] == "Inc" or d.name[:3] == "Exp"):
            row["prev_year_total"] = ""
        else:
            row["prev_year_total"] = prev_year_total

        data.append(row)

    return data


######################################################################################################
## overriden from financial_statements.py
######################################################################################################
def add_total_row(end_month_and_year, out, root_type, balance_must_be, period_list, company_currency):
    total_row = {
        "account_name": _("Total {0} ({1})").format(_(root_type), _(balance_must_be)),
        "account": _("Total {0} ({1})").format(_(root_type), _(balance_must_be)),
        "currency": company_currency,
        "opening_balance": 0.0,
    }

    for row in out:
        if not row.get("parent_account"):
            for period in period_list:
                total_row.setdefault(period.key, 0.0)
                total_row[period.key] += row.get(period.key, 0.0)
                row[period.key] = row.get(period.key, 0.0)

            total_row.setdefault("total", 0.0)
            total_row["total"] += flt(row["total"])
            total_row["opening_balance"] += row["opening_balance"]
            row["total"] = ""

    if "total" in total_row:
        out.append(total_row)
        out.append({}) # blank row after Total


######################################################################################################
##
######################################################################################################
def set_gl_entries_by_account(company, from_date, to_date, root_lft, root_rgt, filters, gl_entries_by_account, ignore_closing_entries=False):
    additional_conditions = get_additional_conditions(from_date, ignore_closing_entries, filters)

    accounts = frappe.db.sql_list(
        """
        SELECT 
            name,
            print_group 
        FROM 
            `tabAccount`
        WHERE 
            lft >= %s
            AND rgt <= %s
            AND company = %s
        """,
        (root_lft, root_rgt, company, ),
    )
    # Returns a dict like { "account": [gl entries], ... }

    if accounts:
        additional_conditions += " AND account IN ({})".format(
            ", ".join(frappe.db.escape(d) for d in accounts)
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
            as_dict=True,
        )  # nosec

        if filters and filters.get("presentation_currency"):
            convert_to_presentation_currency(gl_entries, get_currency(filters), filters.get("company"))

        for entry in gl_entries:
            gl_entries_by_account.setdefault(entry.account, []).append(entry)

        return gl_entries_by_account


