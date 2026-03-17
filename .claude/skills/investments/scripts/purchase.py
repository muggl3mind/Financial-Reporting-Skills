#!/usr/bin/env python3
"""
Securities Purchase Script
Records the purchase of investment securities.
"""

import argparse
import json
import os
import sys
from datetime import datetime
import uuid

# Add shared module to path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_PATH = os.path.join(SCRIPT_DIR, '..', '..', '..', '..', 'lib')
sys.path.insert(0, LIB_PATH)

from utils import (save_journal_csv, save_journal_xlsx, sanitize_filename,
                   OPENPYXL_AVAILABLE, validate_positive_amount, validate_positive_int,
                   validate_date, get_or_create_run_timestamp)
from accounts import SECURITIES_ACCOUNTS, get_securities_account

# Use shared cash account constant
CASH_ACCOUNT = SECURITIES_ACCOUNTS['cash']


def parse_args():
    parser = argparse.ArgumentParser(description='Record securities purchase')
    parser.add_argument('--name', required=True, help='Security name')
    parser.add_argument('--date', required=True, help='Purchase date (YYYY-MM-DD)')
    parser.add_argument('--shares', required=True, type=int, help='Number of shares')
    parser.add_argument('--price', required=True, type=float, help='Price per share')
    parser.add_argument('--security-type', default='AFS',
                        choices=['AFS', 'Trading'],
                        help='Security classification (default: AFS)')
    parser.add_argument('--output-dir', default='data/output/investments',
                        help='Output directory')
    parser.add_argument('--run-timestamp', default=None,
                        help='Timestamp for output folder (YYYY-MM-DD_HH-MM). If not provided, uses current time.')
    return parser.parse_args()


def generate_purchase_journal(name, date, shares, price, security_type):
    """Generate journal entries for securities purchase."""
    total_cost = shares * price
    entries = []
    description = f"Purchase {shares} shares of {name} @ ${price:.2f}"

    # Determine investment account based on security type
    inv_account = get_securities_account(security_type)

    # Debit: Investment account
    entries.append({
        'date': date,
        'type': 'Securities Purchase',
        'asset_name': name,
        'account_code': inv_account['code'],
        'account_name': inv_account['name'],
        'description': description,
        'debit': total_cost,
        'credit': 0.00
    })

    # Credit: Cash
    entries.append({
        'date': date,
        'type': 'Securities Purchase',
        'asset_name': name,
        'account_code': CASH_ACCOUNT['code'],
        'account_name': CASH_ACCOUNT['name'],
        'description': description,
        'debit': 0.00,
        'credit': total_cost
    })

    return entries, {
        'type': 'purchase',
        'security': name,
        'date': date,
        'shares': shares,
        'price_per_share': price,
        'total_cost': total_cost,
        'security_type': security_type
    }


def update_securities_register(transaction_info, register_path):
    """Update securities register JSON file."""
    if os.path.exists(register_path):
        with open(register_path, 'r') as f:
            register = json.load(f)
    else:
        register = {'transactions': [], 'holdings': {}, 'metadata': {'total_transactions': 0}}

    # Add transaction
    transaction_id = f"txn_{str(uuid.uuid4())[:8]}"
    transaction_info['id'] = transaction_id
    transaction_info['recorded_at'] = datetime.now().isoformat()
    register['transactions'].append(transaction_info)

    # Update holdings
    security = transaction_info['security']
    if security not in register['holdings']:
        register['holdings'][security] = {
            'shares': 0,
            'cost_basis': 0,
            'security_type': transaction_info['security_type']
        }

    register['holdings'][security]['shares'] += transaction_info['shares']
    register['holdings'][security]['cost_basis'] += transaction_info['total_cost']

    register['metadata']['total_transactions'] = len(register['transactions'])
    register['metadata']['last_updated'] = datetime.now().isoformat()

    with open(register_path, 'w') as f:
        json.dump(register, f, indent=2)

    return transaction_id


def main():
    args = parse_args()

    # Validate inputs
    try:
        validate_date(args.date, "Purchase date")
        validate_positive_int(args.shares, "Shares")
        validate_positive_amount(args.price, "Price per share")
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # Generate purchase journal entry
    entries, transaction_info = generate_purchase_journal(
        args.name, args.date, args.shares, args.price, args.security_type
    )

    # Create timestamped output folder
    run_timestamp = get_or_create_run_timestamp(args.run_timestamp)
    run_output_dir = os.path.join(args.output_dir, run_timestamp)
    os.makedirs(run_output_dir, exist_ok=True)

    safe_name = sanitize_filename(args.name)

    # Transaction files go to timestamped folder
    csv_path = os.path.join(run_output_dir, f'securities_purchase_{safe_name}_{args.date}.csv')
    xlsx_path = os.path.join(run_output_dir, f'securities_purchase_{safe_name}_{args.date}.xlsx')

    # Register stays at root level
    register_path = os.path.join(args.output_dir, 'securities_register.json')

    save_journal_csv(entries, csv_path)
    xlsx_ok = save_journal_xlsx(entries, xlsx_path, f"Purchase - {args.name}")
    txn_id = update_securities_register(transaction_info, register_path)

    # Determine investment account for display
    inv_account = get_securities_account(args.security_type)

    # Print summary
    print(f"\n{'='*60}")
    print(f"SECURITIES PURCHASE")
    print(f"{'='*60}")
    print(f"\nSecurity: {args.name}")
    print(f"Transaction ID: {txn_id}")
    print(f"Date: {args.date}")
    print(f"Shares: {args.shares}")
    print(f"Price per Share: ${args.price:,.2f}")
    print(f"Total Cost: ${transaction_info['total_cost']:,.2f}")
    print(f"Classification: {args.security_type}")
    print(f"\n--- Journal Entry ---")
    print(f"Dr {inv_account['code']} {inv_account['name']}: ${transaction_info['total_cost']:,.2f}")
    print(f"Cr {CASH_ACCOUNT['code']} {CASH_ACCOUNT['name']}: ${transaction_info['total_cost']:,.2f}")
    print(f"\n--- Output Files ---")
    print(f"Journal (CSV): {csv_path}")
    if xlsx_ok:
        print(f"Journal (XLSX): {xlsx_path}")
    print(f"Securities Register: {register_path}")
    print(f"\n{'='*60}")


if __name__ == '__main__':
    main()
