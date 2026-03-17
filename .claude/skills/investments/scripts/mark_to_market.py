#!/usr/bin/env python3
"""
Mark-to-Market Script
Records period-end fair value adjustments for investment securities.

AFS Securities: Unrealized gain/loss to OCI (Accumulated Other Comprehensive Income)
Trading Securities: Unrealized gain/loss to P&L
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
                   OPENPYXL_AVAILABLE, validate_positive_amount, validate_date,
                   get_or_create_run_timestamp)
from accounts import SECURITIES_ACCOUNTS, get_securities_account


def parse_args():
    parser = argparse.ArgumentParser(description='Record fair value adjustment for securities')
    parser.add_argument('--name', required=True, help='Security name')
    parser.add_argument('--date', required=True, help='Valuation date (YYYY-MM-DD)')
    parser.add_argument('--fair-value', required=True, type=float, help='Current fair market value')
    parser.add_argument('--output-dir', default='data/output/investments',
                        help='Output directory')
    parser.add_argument('--run-timestamp', default=None,
                        help='Timestamp for output folder (YYYY-MM-DD_HH-MM). If not provided, uses current time.')
    return parser.parse_args()


def load_securities_register(output_dir):
    """Load securities register to get current holdings."""
    register_path = os.path.join(output_dir, 'securities_register.json')
    if os.path.exists(register_path):
        with open(register_path, 'r') as f:
            return json.load(f)
    return {'transactions': [], 'holdings': {}, 'fmv_adjustments': [], 'metadata': {}}


def get_current_book_value(register, security_name):
    """Get current book value (cost basis + prior FMV adjustments) for a security."""
    holding = register['holdings'].get(security_name)
    if not holding:
        return None, None, None

    cost_basis = holding['cost_basis']
    security_type = holding.get('security_type', 'AFS')

    # Get prior FMV adjustment balance
    prior_fmv_adjustment = holding.get('fmv_adjustment', 0)
    current_book_value = cost_basis + prior_fmv_adjustment

    return current_book_value, cost_basis, security_type


def generate_fmv_journal(name, date, current_book_value, fair_value, security_type):
    """Generate journal entry for FMV adjustment."""
    entries = []
    adjustment = fair_value - current_book_value

    if adjustment == 0:
        return entries, {'adjustment': 0, 'type': 'none'}

    is_gain = adjustment > 0

    # Determine investment account
    inv_account = get_securities_account(security_type)

    description = f"FMV adjustment - {name}"

    if security_type == 'AFS':
        # AFS: Unrealized gain/loss goes to OCI (equity)
        oci_account = SECURITIES_ACCOUNTS['oci']
        if is_gain:
            # Dr Investment, Cr OCI
            entries.append({
                'date': date,
                'type': 'FMV Adjustment',
                'asset_name': name,
                'account_code': inv_account['code'],
                'account_name': inv_account['name'],
                'description': description,
                'debit': abs(adjustment),
                'credit': 0.00
            })
            entries.append({
                'date': date,
                'type': 'FMV Adjustment',
                'asset_name': name,
                'account_code': oci_account['code'],
                'account_name': oci_account['name'],
                'description': description,
                'debit': 0.00,
                'credit': abs(adjustment)
            })
        else:
            # Dr OCI, Cr Investment
            entries.append({
                'date': date,
                'type': 'FMV Adjustment',
                'asset_name': name,
                'account_code': oci_account['code'],
                'account_name': oci_account['name'],
                'description': description,
                'debit': abs(adjustment),
                'credit': 0.00
            })
            entries.append({
                'date': date,
                'type': 'FMV Adjustment',
                'asset_name': name,
                'account_code': inv_account['code'],
                'account_name': inv_account['name'],
                'description': description,
                'debit': 0.00,
                'credit': abs(adjustment)
            })
    else:
        # Trading: Unrealized gain/loss goes to P&L
        if is_gain:
            # Dr Investment, Cr Unrealized Gain
            entries.append({
                'date': date,
                'type': 'FMV Adjustment',
                'asset_name': name,
                'account_code': inv_account['code'],
                'account_name': inv_account['name'],
                'description': description,
                'debit': abs(adjustment),
                'credit': 0.00
            })
            entries.append({
                'date': date,
                'type': 'FMV Adjustment',
                'asset_name': name,
                'account_code': SECURITIES_ACCOUNTS['unrealized_gain']['code'],
                'account_name': SECURITIES_ACCOUNTS['unrealized_gain']['name'],
                'description': description,
                'debit': 0.00,
                'credit': abs(adjustment)
            })
        else:
            # Dr Unrealized Loss, Cr Investment
            entries.append({
                'date': date,
                'type': 'FMV Adjustment',
                'asset_name': name,
                'account_code': SECURITIES_ACCOUNTS['unrealized_loss']['code'],
                'account_name': SECURITIES_ACCOUNTS['unrealized_loss']['name'],
                'description': description,
                'debit': abs(adjustment),
                'credit': 0.00
            })
            entries.append({
                'date': date,
                'type': 'FMV Adjustment',
                'asset_name': name,
                'account_code': inv_account['code'],
                'account_name': inv_account['name'],
                'description': description,
                'debit': 0.00,
                'credit': abs(adjustment)
            })

    return entries, {
        'type': 'fmv_adjustment',
        'security': name,
        'date': date,
        'prior_book_value': current_book_value,
        'fair_value': fair_value,
        'adjustment': adjustment,
        'is_gain': is_gain,
        'security_type': security_type
    }


def update_securities_register(register, transaction_info, output_dir):
    """Update securities register with FMV adjustment."""
    register_path = os.path.join(output_dir, 'securities_register.json')

    # Initialize fmv_adjustments if not present
    if 'fmv_adjustments' not in register:
        register['fmv_adjustments'] = []

    # Add adjustment transaction
    transaction_id = f"fmv_{str(uuid.uuid4())[:8]}"
    transaction_info['id'] = transaction_id
    transaction_info['recorded_at'] = datetime.now().isoformat()
    register['fmv_adjustments'].append(transaction_info)

    # Update holding's FMV adjustment balance
    security = transaction_info['security']
    if security in register['holdings']:
        current_adj = register['holdings'][security].get('fmv_adjustment', 0)
        register['holdings'][security]['fmv_adjustment'] = current_adj + transaction_info['adjustment']
        register['holdings'][security]['last_fmv'] = transaction_info['fair_value']
        register['holdings'][security]['last_fmv_date'] = transaction_info['date']

    register['metadata']['last_updated'] = datetime.now().isoformat()

    with open(register_path, 'w') as f:
        json.dump(register, f, indent=2)

    return transaction_id


def main():
    args = parse_args()

    # Validate inputs
    try:
        validate_date(args.date, "Valuation date")
        validate_positive_amount(args.fair_value, "Fair value", allow_zero=True)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # Load register and get current book value
    register = load_securities_register(args.output_dir)
    current_book_value, cost_basis, security_type = get_current_book_value(register, args.name)

    if current_book_value is None:
        print(f"Error: Security '{args.name}' not found in holdings")
        return

    # Generate FMV adjustment journal entry
    entries, transaction_info = generate_fmv_journal(
        args.name, args.date, current_book_value, args.fair_value, security_type
    )

    if transaction_info['adjustment'] == 0:
        print(f"\nNo adjustment needed - fair value equals book value")
        return

    # Create timestamped output folder
    run_timestamp = get_or_create_run_timestamp(args.run_timestamp)
    run_output_dir = os.path.join(args.output_dir, run_timestamp)
    os.makedirs(run_output_dir, exist_ok=True)

    safe_name = sanitize_filename(args.name)

    # Transaction files go to timestamped folder
    csv_path = os.path.join(run_output_dir, f'fmv_adjustment_{safe_name}_{args.date}.csv')
    xlsx_path = os.path.join(run_output_dir, f'fmv_adjustment_{safe_name}_{args.date}.xlsx')

    save_journal_csv(entries, csv_path)
    xlsx_ok = save_journal_xlsx(entries, xlsx_path, f"FMV Adjustment - {args.name}")

    # Register stays at root level
    txn_id = update_securities_register(register, transaction_info, args.output_dir)

    # Print summary
    print(f"\n{'='*60}")
    print(f"FAIR VALUE ADJUSTMENT")
    print(f"{'='*60}")
    print(f"\nSecurity: {args.name}")
    print(f"Transaction ID: {txn_id}")
    print(f"Date: {args.date}")
    print(f"Classification: {security_type}")
    print(f"\n--- Valuation ---")
    print(f"Cost Basis: ${cost_basis:,.2f}")
    print(f"Prior Book Value: ${current_book_value:,.2f}")
    print(f"Fair Value: ${args.fair_value:,.2f}")
    print(f"\n--- Adjustment ---")
    if transaction_info['is_gain']:
        print(f"UNREALIZED GAIN: ${transaction_info['adjustment']:,.2f}")
    else:
        print(f"UNREALIZED LOSS: ${abs(transaction_info['adjustment']):,.2f}")

    if security_type == 'AFS':
        print(f"Recorded to: OCI (Equity)")
    else:
        print(f"Recorded to: Income Statement")

    print(f"\n--- Output Files ---")
    print(f"Journal (CSV): {csv_path}")
    if xlsx_ok:
        print(f"Journal (XLSX): {xlsx_path}")
    print(f"\n{'='*60}")


if __name__ == '__main__':
    main()
