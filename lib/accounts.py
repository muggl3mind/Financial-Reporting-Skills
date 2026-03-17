#!/usr/bin/env python3
"""
Chart of Accounts - Single Source of Truth

This module defines all financial statement accounts used across the project.
All modules (fixed-assets, investments, capital-accounts) import from here.
"""

# =============================================================================
# CHART OF ACCOUNTS - Complete list of all GL accounts
# =============================================================================

CHART_OF_ACCOUNTS = {
    # ===== ASSETS (1000-1899) =====
    '1000': 'Cash and Cash Equivalents',
    '1100': 'Accounts Receivable',
    '1200': 'Inventory',
    '1300': 'Prepaid Expenses',
    '1400': 'Other Current Assets',
    '1500': 'Property Plant & Equipment',
    '1600': 'Accumulated Depreciation',
    '1650': 'Accumulated Amortization',
    '1700': 'Intangible Assets',
    '1750': 'Investments - Available for Sale',
    '1760': 'Investments - Trading Securities',
    '1800': 'Other Non-Current Assets',

    # ===== LIABILITIES (2000-2899) =====
    '2000': 'Accounts Payable',
    '2100': 'Accrued Expenses',
    '2200': 'Accrued Payroll',
    '2300': 'Short-term Debt',
    '2400': 'Deferred Revenue',
    '2500': 'Other Current Liabilities',
    '2600': 'Long-term Debt',
    '2700': 'Other Non-Current Liabilities',

    # ===== EQUITY (3000-3899) =====
    '3000': 'Common Stock',
    '3100': 'Preferred Stock',
    '3200': 'Additional Paid-in Capital',
    '3300': 'PE Capital Contributions',
    '3400': 'Retained Earnings',
    '3500': 'Current Year Earnings',
    '3600': 'Accumulated Other Comprehensive Income',

    # ===== REVENUE (4000-4899) =====
    '4000': 'Product Revenue',
    '4100': 'Service Revenue',
    '4200': 'Licensing/Royalty Revenue',
    '4300': 'Grant Revenue',
    '4900': 'Other Operating Revenue',

    # ===== COST OF GOODS SOLD (5000-5899) =====
    '5000': 'Cost of Goods Sold',

    # ===== OPERATING EXPENSES (6000-6899) =====
    '6100': 'Salaries and Wages Expense',
    '6150': 'Benefits and Payroll Tax Expense',
    '6200': 'Rent Expense',
    '6250': 'Utilities Expense',
    '6300': 'Depreciation Expense',
    '6350': 'Amortization Expense',
    '6400': 'Office Expenses',
    '6450': 'Software and IT Expense',
    '6500': 'Professional Services Expense',
    '6550': 'Travel and Entertainment Expense',
    '6600': 'Marketing and Advertising Expense',
    '6650': 'Research and Development Expense',
    '6700': 'Insurance Expense',
    '6750': 'PE Management Fees',
    '6900': 'Other Operating Expenses',

    # ===== OTHER INCOME (7000-7899) =====
    '7000': 'Interest Income',
    '7100': 'Dividend Income',
    '7200': 'Gain on Disposal of Fixed Assets',
    '7300': 'Realized Gain on Sale of Investments',
    '7400': 'Unrealized Gain (Trading Securities)',
    '7900': 'Other Non-Operating Income',

    # ===== OTHER EXPENSES (8000-8899) =====
    '8000': 'Interest Expense',
    '8100': 'Loss on Disposal of Fixed Assets',
    '8200': 'Realized Loss on Sale of Investments',
    '8300': 'Unrealized Loss (Trading Securities)',
    '8800': 'Income Tax Expense',
    '8900': 'Other Non-Operating Expenses',
}


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_account(code):
    """Get account dict by code.

    Args:
        code: Account code (e.g., '1000', '2000')

    Returns:
        Dictionary with 'code' and 'name' keys

    Raises:
        KeyError: If account code not found in CHART_OF_ACCOUNTS
    """
    return {'code': code, 'name': CHART_OF_ACCOUNTS[code]}


# =============================================================================
# DEFAULT USEFUL LIVES - Based on IRS MACRS Guidelines
# =============================================================================

DEFAULT_USEFUL_LIVES = {
    'Equipment': 7,   # 7-year MACRS property
    'Vehicle': 5,     # 5-year MACRS property
    'Furniture': 7,   # 7-year MACRS property
    'Computer': 5,    # 5-year MACRS property
    'Building': 39,   # Nonresidential real property
    'Software': 3,    # 3-year MACRS property
    'Patent': 15,     # 15-year property (intangibles)
}


def get_default_useful_life(category):
    """Get default useful life for an asset category.

    Args:
        category: Asset category (Equipment, Vehicle, Furniture, Computer,
                  Building, Software, Patent)

    Returns:
        Integer useful life in years

    Raises:
        ValueError: If category not recognized
    """
    if category not in DEFAULT_USEFUL_LIVES:
        valid = ', '.join(DEFAULT_USEFUL_LIVES.keys())
        raise ValueError(f"Unknown category '{category}'. Valid categories: {valid}")
    return DEFAULT_USEFUL_LIVES[category]


# =============================================================================
# CONVENIENCE MAPPINGS - Named references for common account groups
# =============================================================================

# Securities-related accounts
SECURITIES_ACCOUNTS = {
    'cash': get_account('1000'),
    'investments_afs': get_account('1750'),
    'investments_trading': get_account('1760'),
    'realized_gain': get_account('7300'),
    'realized_loss': get_account('8200'),
    'unrealized_gain': get_account('7400'),
    'unrealized_loss': get_account('8300'),
    'oci': get_account('3600'),
}


def get_securities_account(security_type):
    """Get the investment account for a security type.

    Args:
        security_type: 'AFS' or 'Trading'

    Returns:
        Dictionary with code and name
    """
    if security_type == 'Trading':
        return SECURITIES_ACCOUNTS['investments_trading']
    else:
        return SECURITIES_ACCOUNTS['investments_afs']


def get_account_mapping(category):
    """Map asset category to account codes.

    Args:
        category: Asset category (Equipment, Vehicle, Furniture, Computer,
                  Building, Software, Patent)

    Returns:
        Dictionary with account codes and names for:
        - asset_code/asset_name: Balance sheet asset account
        - accum_code/accum_name: Accumulated depreciation/amortization
        - expense_code/expense_name: Depreciation/amortization expense
        - gain_code/gain_name: Disposal gain account
        - loss_code/loss_name: Disposal loss account
    """
    intangible_categories = ['Software', 'Patent']

    if category in intangible_categories:
        asset = get_account('1700')
        accum = get_account('1650')
        expense = get_account('6350')
    else:
        asset = get_account('1500')
        accum = get_account('1600')
        expense = get_account('6300')

    gain = get_account('7200')
    loss = get_account('8100')

    return {
        'asset_code': asset['code'],
        'asset_name': asset['name'],
        'accum_code': accum['code'],
        'accum_name': accum['name'],
        'expense_code': expense['code'],
        'expense_name': expense['name'],
        'gain_code': gain['code'],
        'gain_name': gain['name'],
        'loss_code': loss['code'],
        'loss_name': loss['name'],
    }
