# Copyright (c) 2022, Farabi Hussain

import frappe
from frappe import _, scrub
from erpnext.accounts.report.financial_statements import *


def execute(filters = None):
    period_list = get_period_list(filters.to_fiscal_year, filters.period_start_date, filters.period_end_date, filters.filter_based_on, filters.periodicity, filters.period_end_month, company = filters.company,)
    income  = get_data(filters.period_end_month, filters.to_fiscal_year, filters.company, "Income", "Credit", period_list, filters = filters, accumulated_values = filters.accumulated_values, ignore_closing_entries = True, ignore_accumulated_values_for_fy = True,)
    expense = get_data(filters.period_end_month, filters.to_fiscal_year, filters.company, "Expense", "Debit", period_list, filters = filters, accumulated_values = filters.accumulated_values, ignore_closing_entries = True, ignore_accumulated_values_for_fy = True,)

    data = []
    data.extend(income or [])
    data.extend(expense or [])

    columns = get_columns(filters.period_end_month, filters.to_fiscal_year, filters.periodicity, period_list, filters.accumulated_values, filters.company)
    return columns, data, None, None, None


## overriden function from financial_statements.py
# goes through the columns and appends them based on the user-selected month
def get_columns(period_end_month, period_end_year, periodicity, period_list, accumulated_values=1, company=None):
    columns = [{"fieldname": "account", "label": _("Account"), "fieldtype": "Link", "options": "Account", "width": 200,}]
    end_month_and_year = (period_end_month[0:3] + " " + period_end_year)

    for period in period_list:
        columns.append({"fieldname": period.key, "label": period.label, "fieldtype": "Currency", "options": "currency", "width": 155,})
        if (period.label == end_month_and_year): break

    columns.append({"fieldname": "total", "label": _("Total"), "fieldtype": "Currency", "width": 150,})
    # columns.append({"fieldname": "totals", "label": _("Totals"), "fieldtype": "Currency", "width": 150,})

    return columns


## overriden function from financial_statements.py
# 
def get_period_list(to_fiscal_year, period_start_date, period_end_date, filter_based_on, periodicity, period_end_month, accumulated_values=False, company=None, reset_period_on_fy_change=True, ignore_fiscal_year=False):
    # Get a list of dict {"from_date": from_date, "to_date": to_date, "key": key, "label": label}
    # Periodicity can be (Yearly, Quarterly, Monthly)

    from_fiscal_year = to_fiscal_year
    build_start_year_and_date = (str(int(from_fiscal_year)-1) + '-01-01')

    fiscal_year = get_fiscal_year_data(from_fiscal_year, to_fiscal_year)
    validate_fiscal_year(fiscal_year, from_fiscal_year, to_fiscal_year)
    # year_start_date = getdate(fiscal_year.year_start_date)
    year_start_date = getdate(build_start_year_and_date)
    year_end_date = getdate(fiscal_year.year_end_date)

    months_to_add = {"Yearly": 12, "Half-Yearly": 6, "Quarterly": 3, "Monthly": 1}[periodicity]
    period_list = []

    start_date = year_start_date
    months = get_months(year_start_date, year_end_date)

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


## overriden function from financial_statements.py
# 
def get_data(period_end_month, period_end_year, company, root_type, balance_must_be, period_list, filters=None, accumulated_values=1, only_current_fiscal_year=True, ignore_closing_entries=False, ignore_accumulated_values_for_fy=False, total=True,):
    end_month_and_year = (period_end_month[0:3] + " " + period_end_year)
    accounts = get_accounts(company, root_type)
    if not accounts:
        return None

    accounts, accounts_by_name, parent_children_map = filter_accounts(accounts)
    company_currency = get_appropriate_currency(company, filters)
    gl_entries_by_account = {}

    for root in frappe.db.sql(
        """
        SELECT 
            lft, rgt 
        FROM 
            tabAccount
        WHERE 
            root_type=%s 
            AND ifnull(parent_account, '') = ''
        """,
        root_type,
        as_dict=1,
    ):

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

    calculate_values(accounts_by_name, gl_entries_by_account, period_list, accumulated_values, ignore_accumulated_values_for_fy)
    accumulate_values_into_parents(accounts, accounts_by_name, period_list)
    out = prepare_data(end_month_and_year, accounts, balance_must_be, period_list, company_currency)
    out = filter_out_zero_value_rows(out, parent_children_map)

    if out and total:
        add_total_row(end_month_and_year, out, root_type, balance_must_be, period_list, company_currency)

    return out

## overriden function from financial_statements.py
# 
def prepare_data(end_month_and_year, accounts, balance_must_be, period_list, company_currency):
    data = []
    year_start_date = period_list[0]["year_start_date"].strftime("%Y-%m-%d")
    year_end_date = period_list[-1]["year_end_date"].strftime("%Y-%m-%d")

    for d in accounts:
        # add to output
        has_value = False
        total = 0
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
            if abs(row[period.key]) >= 0.005:
                has_value = True
                total += flt(row[period.key])
            
            if (period.label == end_month_and_year):
                break

        row["has_value"] = has_value
        row["total"] = total
        data.append(row)

    return data


## overriden function from financial_statements.py
#
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