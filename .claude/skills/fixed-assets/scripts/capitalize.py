#!/usr/bin/env python3
"""
Fixed Asset Capitalization Script
Records the initial purchase/capitalization of a fixed asset.
Generates journal entry: Dr PP&E/Intangibles, Cr Cash or AP
"""

import argparse
import json
import os
import sys
from datetime import datetime

# Add shared module to path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_PATH = os.path.join(SCRIPT_DIR, '..', '..', '..', '..', 'lib')
sys.path.insert(0, LIB_PATH)
sys.path.insert(0, SCRIPT_DIR)

from accounts import get_account_mapping, SECURITIES_ACCOUNTS, get_account
from utils import (save_journal_csv, save_journal_xlsx, sanitize_filename,
                   OPENPYXL_AVAILABLE, validate_positive_amount, validate_date,
                   get_or_create_run_timestamp)

# Credit account options - all from shared chart of accounts
CREDIT_ACCOUNTS = {
    'cash': SECURITIES_ACCOUNTS['cash'],
    'ap': get_account('2000'),
}


def parse_args():
    parser = argparse.ArgumentParser(description='Record fixed asset purchase/capitalization')
    parser.add_argument('--name', required=True, help='Asset name')
    parser.add_argument('--category', required=True,
                        choices=['Equipment', 'Vehicle', 'Furniture', 'Computer',
                                 'Building', 'Software', 'Patent'],
                        help='Asset category')
    parser.add_argument('--date', required=True, help='Purchase date (YYYY-MM-DD)')
    parser.add_argument('--cost', required=True, type=float, help='Purchase cost')
    parser.add_argument('--credit-account', default='cash',
                        choices=['cash', 'ap'],
                        help='Credit account: cash (paid) or ap (on account)')
    parser.add_argument('--output-dir', default='data/output/fixed-assets',
                        help='Output directory')
    parser.add_argument('--run-timestamp', default=None,
                        help='Timestamp for output folder (YYYY-MM-DD_HH-MM). If not provided, uses current time.')
    return parser.parse_args()


def generate_purchase_journal(asset_name, category, date, cost, credit_account):
    """Generate journal entry for asset purchase."""
    entries = []

    # Get asset account based on category
    accounts = get_account_mapping(category)
    credit_acct = CREDIT_ACCOUNTS[credit_account]

    description = f"Purchase of {asset_name}"

    # Debit: Asset account (PP&E or Intangibles)
    entries.append({
        'date': date,
        'type': 'Capitalization',
        'asset_name': asset_name,
        'account_code': accounts['asset_code'],
        'account_name': accounts['asset_name'],
        'description': description,
        'debit': cost,
        'credit': 0.00
    })

    # Credit: Cash or AP
    entries.append({
        'date': date,
        'type': 'Capitalization',
        'asset_name': asset_name,
        'account_code': credit_acct['code'],
        'account_name': credit_acct['name'],
        'description': description,
        'debit': 0.00,
        'credit': cost
    })

    return entries


def update_asset_register_purchase(output_dir, asset_name, purchase_date, cost, category):
    """Update asset register with purchase information."""
    register_path = os.path.join(output_dir, 'asset_register.json')

    if os.path.exists(register_path):
        with open(register_path, 'r') as f:
            register = json.load(f)
    else:
        register = {'assets': [], 'metadata': {}}

    # Find the asset and mark as capitalized
    asset_found = False
    for asset in register.get('assets', []):
        if asset['name'] == asset_name:
            asset['purchase_recorded'] = True
            asset['purchase_journal_date'] = purchase_date
            asset_found = True
            break

    # If asset not in register, add basic entry
    if not asset_found:
        register['assets'].append({
            'name': asset_name,
            'category': category,
            'purchase_date': purchase_date,
            'cost': cost,
            'purchase_recorded': True,
            'purchase_journal_date': purchase_date
        })

    register['metadata']['last_updated'] = datetime.now().isoformat()

    with open(register_path, 'w') as f:
        json.dump(register, f, indent=2)


def main():
    args = parse_args()

    # Validate inputs
    try:
        validate_date(args.date, "Purchase date")
        validate_positive_amount(args.cost, "Cost")
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # Generate purchase journal entry
    entries = generate_purchase_journal(
        args.name, args.category, args.date, args.cost, args.credit_account
    )

    # Create timestamped output folder
    run_timestamp = get_or_create_run_timestamp(args.run_timestamp)
    run_output_dir = os.path.join(args.output_dir, run_timestamp)
    os.makedirs(run_output_dir, exist_ok=True)

    safe_name = sanitize_filename(args.name)

    # Transaction files go to timestamped folder
    csv_path = os.path.join(run_output_dir, f'capitalize_{safe_name}.csv')
    xlsx_path = os.path.join(run_output_dir, f'capitalize_{safe_name}.xlsx')

    save_journal_csv(entries, csv_path)
    xlsx_ok = save_journal_xlsx(entries, xlsx_path, f"Capitalization - {args.name}")

    # Update asset register (stays at root level)
    update_asset_register_purchase(args.output_dir, args.name, args.date, args.cost, args.category)

    # Get account info for display
    accounts = get_account_mapping(args.category)
    credit_acct = CREDIT_ACCOUNTS[args.credit_account]

    # Print summary
    print(f"\n{'='*60}")
    print(f"FIXED ASSET CAPITALIZATION")
    print(f"{'='*60}")
    print(f"\nAsset: {args.name}")
    print(f"Category: {args.category}")
    print(f"Purchase Date: {args.date}")
    print(f"Cost: ${args.cost:,.2f}")
    print(f"\n--- Journal Entry ---")
    print(f"Dr {accounts['asset_code']} {accounts['asset_name']}: ${args.cost:,.2f}")
    print(f"Cr {credit_acct['code']} {credit_acct['name']}: ${args.cost:,.2f}")
    print(f"\n--- Output Files ---")
    print(f"Journal (CSV): {csv_path}")
    if xlsx_ok:
        print(f"Journal (XLSX): {xlsx_path}")
    print(f"\n{'='*60}")


if __name__ == '__main__':
    main()
