#!/usr/bin/env python3
"""
Consolidated Journal Generator
Generates a single journal containing ALL entries through a specified date:
- Fixed asset capitalization (purchase) entries
- Depreciation entries (from asset schedules)
- Disposal entries
- Securities purchase/sale entries
- FMV adjustment entries
"""

import argparse
import json
import os
import sys
from datetime import datetime

# Add shared module to path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from utils import save_journal_csv, save_journal_xlsx, OPENPYXL_AVAILABLE, calculate_period_begin, VALID_PERIOD_TYPES
from accounts import get_account_mapping, SECURITIES_ACCOUNTS, get_securities_account

# Add fixed-assets module to path for depreciation calculations
FIXED_ASSETS_PATH = os.path.join(SCRIPT_DIR, '..', 'fixed-assets', 'scripts')
sys.path.insert(0, FIXED_ASSETS_PATH)

from depreciation import calculate_straight_line, generate_schedule, generate_journal_entries


def parse_args():
    parser = argparse.ArgumentParser(description='Generate consolidated journal through a date')
    parser.add_argument('--through-date', required=True,
                        help='Include all entries through this date (YYYY-MM-DD)')
    parser.add_argument('--period-begin', required=False,
                        help='Period begin date (YYYY-MM-DD). Only entries on/after this date.')
    parser.add_argument('--period-type', required=False, choices=VALID_PERIOD_TYPES,
                        help='Period type - auto-calculates period-begin from through-date.')
    parser.add_argument('--fixed-assets-dir', required=False,
                        default='data/output/fixed-assets',
                        help='Directory containing asset_register.json')
    parser.add_argument('--investments-dir', required=False,
                        default='data/output/investments',
                        help='Directory containing securities_register.json')
    parser.add_argument('--output-dir', required=True, help='Output directory')
    return parser.parse_args()


def load_asset_register(assets_dir):
    """Load asset register to get all fixed assets, merging duplicate entries by name."""
    register_path = os.path.join(assets_dir, 'asset_register.json')
    if os.path.exists(register_path):
        with open(register_path, 'r') as f:
            data = json.load(f)

        # Merge duplicate assets by name
        merged = {}
        for asset in data.get('assets', []):
            name = asset['name']
            if name in merged:
                # Merge properties from both entries
                merged[name].update(asset)
            else:
                merged[name] = asset.copy()

        return {'assets': list(merged.values())}
    return {'assets': []}


def load_securities_register(investments_dir):
    """Load securities register to get all securities transactions."""
    register_path = os.path.join(investments_dir, 'securities_register.json')
    if os.path.exists(register_path):
        with open(register_path, 'r') as f:
            return json.load(f)
    return {'transactions': [], 'fmv_adjustments': []}


def generate_capitalization_entries(asset, through_date):
    """Generate purchase/capitalization journal entry for a fixed asset."""
    entries = []
    through_dt = datetime.strptime(through_date, '%Y-%m-%d')
    purchase_dt = datetime.strptime(asset['purchase_date'], '%Y-%m-%d')

    # Only include if purchase is within date range and recorded
    if purchase_dt > through_dt:
        return entries

    if not asset.get('purchase_recorded', False):
        return entries

    accounts = get_account_mapping(asset['category'])
    description = f"Purchase of {asset['name']}"

    # Debit: Asset account (PP&E or Intangibles)
    entries.append({
        'date': asset['purchase_date'],
        'type': 'Capitalization',
        'asset_name': asset['name'],
        'account_code': accounts['asset_code'],
        'account_name': accounts['asset_name'],
        'description': description,
        'debit': asset['cost'],
        'credit': 0.00
    })

    # Credit: Cash (default assumption)
    cash = SECURITIES_ACCOUNTS['cash']
    entries.append({
        'date': asset['purchase_date'],
        'type': 'Capitalization',
        'asset_name': asset['name'],
        'account_code': cash['code'],
        'account_name': cash['name'],
        'description': description,
        'debit': 0.00,
        'credit': asset['cost']
    })

    return entries


def generate_depreciation_entries(asset, through_date):
    """Generate all depreciation journal entries for an asset through the specified date."""
    entries = []

    # Skip if asset doesn't have depreciation info
    if 'salvage_value' not in asset or 'useful_life_years' not in asset or 'reporting_period' not in asset:
        return entries

    through_dt = datetime.strptime(through_date, '%Y-%m-%d')

    # Get account mapping
    accounts = get_account_mapping(asset['category'])

    # Calculate depreciation
    calc_results = calculate_straight_line(
        asset['cost'],
        asset['salvage_value'],
        asset['useful_life_years'],
        asset['reporting_period']
    )

    # Generate schedule (pass disposal_date if asset was disposed)
    disposal_date = asset.get('disposal_date')
    schedule = generate_schedule(
        asset['name'],
        asset['purchase_date'],
        asset['cost'],
        asset['salvage_value'],
        asset['useful_life_years'],
        asset['reporting_period'],
        calc_results,
        disposal_date
    )

    # Generate journal entries for periods through the date
    period_label = {
        'monthly': 'Monthly',
        'quarterly': 'Quarterly',
        'annual': 'Annual'
    }

    for period in schedule:
        period_end = datetime.strptime(period['end_date'], '%Y-%m-%d')

        # Skip periods after through_date
        if period_end > through_dt:
            break

        # Determine if this is the final (disposal) period
        is_disposal_period = False
        if disposal_date:
            disposal_dt = datetime.strptime(disposal_date, '%Y-%m-%d')
            # Schedule stops at disposal date, so last period ends at disposal
            if period_end == disposal_dt:
                is_disposal_period = True

        if is_disposal_period:
            description = f"Depreciation through disposal - {asset['name']}"
        else:
            description = f"{period_label[asset['reporting_period']]} depreciation - {asset['name']}"

        # Debit expense
        entries.append({
            'date': period['end_date'],
            'type': 'Depreciation',
            'asset_name': asset['name'],
            'account_code': accounts['expense_code'],
            'account_name': accounts['expense_name'],
            'description': description,
            'debit': period['depreciation'],
            'credit': 0.00
        })

        # Credit accumulated depreciation
        entries.append({
            'date': period['end_date'],
            'type': 'Depreciation',
            'asset_name': asset['name'],
            'account_code': accounts['accum_code'],
            'account_name': accounts['accum_name'],
            'description': description,
            'debit': 0.00,
            'credit': period['depreciation']
        })

    return entries


def generate_disposal_entries_from_asset(asset, through_date):
    """Generate disposal journal entries if asset was disposed through the date."""
    entries = []

    if 'disposal_date' not in asset:
        return entries

    through_dt = datetime.strptime(through_date, '%Y-%m-%d')
    disposal_dt = datetime.strptime(asset['disposal_date'], '%Y-%m-%d')

    if disposal_dt > through_dt:
        return entries

    # Get account mapping
    accounts = get_account_mapping(asset['category'])

    # Get disposal info from asset
    disposal_info = asset.get('disposal_info', {})
    proceeds = disposal_info.get('proceeds', 0)
    accumulated_depr = disposal_info.get('accumulated_depreciation', 0)
    gain_loss = disposal_info.get('gain_loss', 0)
    is_gain = gain_loss >= 0

    description = f"Disposal of {asset['name']}"
    cash = SECURITIES_ACCOUNTS['cash']

    # Debit Cash
    entries.append({
        'date': asset['disposal_date'],
        'type': 'Disposal',
        'asset_name': asset['name'],
        'account_code': cash['code'],
        'account_name': cash['name'],
        'description': description,
        'debit': proceeds,
        'credit': 0.00
    })

    # Debit Accumulated Depreciation
    entries.append({
        'date': asset['disposal_date'],
        'type': 'Disposal',
        'asset_name': asset['name'],
        'account_code': accounts['accum_code'],
        'account_name': accounts['accum_name'],
        'description': description,
        'debit': accumulated_depr,
        'credit': 0.00
    })

    # If loss, debit loss account
    if not is_gain and gain_loss != 0:
        entries.append({
            'date': asset['disposal_date'],
            'type': 'Disposal',
            'asset_name': asset['name'],
            'account_code': accounts['loss_code'],
            'account_name': accounts['loss_name'],
            'description': description,
            'debit': abs(gain_loss),
            'credit': 0.00
        })

    # Credit Asset
    entries.append({
        'date': asset['disposal_date'],
        'type': 'Disposal',
        'asset_name': asset['name'],
        'account_code': accounts['asset_code'],
        'account_name': accounts['asset_name'],
        'description': description,
        'debit': 0.00,
        'credit': asset['cost']
    })

    # If gain, credit gain account
    if is_gain and gain_loss > 0:
        entries.append({
            'date': asset['disposal_date'],
            'type': 'Disposal',
            'asset_name': asset['name'],
            'account_code': accounts['gain_code'],
            'account_name': accounts['gain_name'],
            'description': description,
            'debit': 0.00,
            'credit': gain_loss
        })

    return entries


def generate_securities_entries(transaction, through_date):
    """Generate journal entries for a securities transaction."""
    entries = []
    through_dt = datetime.strptime(through_date, '%Y-%m-%d')
    txn_dt = datetime.strptime(transaction['date'], '%Y-%m-%d')

    if txn_dt > through_dt:
        return entries

    security_type = transaction.get('security_type', 'AFS')
    inv_account = get_securities_account(security_type)
    cash = SECURITIES_ACCOUNTS['cash']

    if transaction['type'] == 'purchase':
        description = f"Purchase {transaction['shares']} shares of {transaction['security']} @ ${transaction['price_per_share']:.2f}"

        # Debit Investment
        entries.append({
            'date': transaction['date'],
            'type': 'Securities Purchase',
            'asset_name': transaction['security'],
            'account_code': inv_account['code'],
            'account_name': inv_account['name'],
            'description': description,
            'debit': transaction['total_cost'],
            'credit': 0.00
        })

        # Credit Cash
        entries.append({
            'date': transaction['date'],
            'type': 'Securities Purchase',
            'asset_name': transaction['security'],
            'account_code': cash['code'],
            'account_name': cash['name'],
            'description': description,
            'debit': 0.00,
            'credit': transaction['total_cost']
        })

    elif transaction['type'] == 'sale':
        description = f"Sale of {transaction['shares']} shares of {transaction['security']}"
        gain_loss = transaction.get('gain_loss', 0)
        is_gain = gain_loss >= 0
        realized_gain = SECURITIES_ACCOUNTS['realized_gain']
        realized_loss = SECURITIES_ACCOUNTS['realized_loss']

        # Debit Cash
        entries.append({
            'date': transaction['date'],
            'type': 'Securities Sale',
            'asset_name': transaction['security'],
            'account_code': cash['code'],
            'account_name': cash['name'],
            'description': description,
            'debit': transaction['proceeds'],
            'credit': 0.00
        })

        # If loss, debit loss account
        if not is_gain:
            entries.append({
                'date': transaction['date'],
                'type': 'Securities Sale',
                'asset_name': transaction['security'],
                'account_code': realized_loss['code'],
                'account_name': realized_loss['name'],
                'description': description,
                'debit': abs(gain_loss),
                'credit': 0.00
            })

        # Credit Investment
        entries.append({
            'date': transaction['date'],
            'type': 'Securities Sale',
            'asset_name': transaction['security'],
            'account_code': inv_account['code'],
            'account_name': inv_account['name'],
            'description': description,
            'debit': 0.00,
            'credit': transaction['cost_basis']
        })

        # If gain, credit gain account
        if is_gain and gain_loss > 0:
            entries.append({
                'date': transaction['date'],
                'type': 'Securities Sale',
                'asset_name': transaction['security'],
                'account_code': realized_gain['code'],
                'account_name': realized_gain['name'],
                'description': description,
                'debit': 0.00,
                'credit': gain_loss
            })

    return entries


def generate_fmv_entries(adjustment, through_date):
    """Generate journal entries for an FMV adjustment."""
    entries = []
    through_dt = datetime.strptime(through_date, '%Y-%m-%d')
    adj_dt = datetime.strptime(adjustment['date'], '%Y-%m-%d')

    if adj_dt > through_dt:
        return entries

    security_type = adjustment.get('security_type', 'AFS')
    inv_account = get_securities_account(security_type)
    oci = SECURITIES_ACCOUNTS['oci']
    unrealized_gain = SECURITIES_ACCOUNTS['unrealized_gain']
    unrealized_loss = SECURITIES_ACCOUNTS['unrealized_loss']

    adj_amount = adjustment['adjustment']
    is_gain = adj_amount > 0
    description = f"FMV adjustment - {adjustment['security']}"

    if security_type == 'AFS':
        # AFS: Unrealized gain/loss to OCI
        if is_gain:
            entries.append({
                'date': adjustment['date'],
                'type': 'FMV Adjustment',
                'asset_name': adjustment['security'],
                'account_code': inv_account['code'],
                'account_name': inv_account['name'],
                'description': description,
                'debit': abs(adj_amount),
                'credit': 0.00
            })
            entries.append({
                'date': adjustment['date'],
                'type': 'FMV Adjustment',
                'asset_name': adjustment['security'],
                'account_code': oci['code'],
                'account_name': oci['name'],
                'description': description,
                'debit': 0.00,
                'credit': abs(adj_amount)
            })
        else:
            entries.append({
                'date': adjustment['date'],
                'type': 'FMV Adjustment',
                'asset_name': adjustment['security'],
                'account_code': oci['code'],
                'account_name': oci['name'],
                'description': description,
                'debit': abs(adj_amount),
                'credit': 0.00
            })
            entries.append({
                'date': adjustment['date'],
                'type': 'FMV Adjustment',
                'asset_name': adjustment['security'],
                'account_code': inv_account['code'],
                'account_name': inv_account['name'],
                'description': description,
                'debit': 0.00,
                'credit': abs(adj_amount)
            })
    else:
        # Trading: Unrealized gain/loss to P&L
        if is_gain:
            entries.append({
                'date': adjustment['date'],
                'type': 'FMV Adjustment',
                'asset_name': adjustment['security'],
                'account_code': inv_account['code'],
                'account_name': inv_account['name'],
                'description': description,
                'debit': abs(adj_amount),
                'credit': 0.00
            })
            entries.append({
                'date': adjustment['date'],
                'type': 'FMV Adjustment',
                'asset_name': adjustment['security'],
                'account_code': unrealized_gain['code'],
                'account_name': unrealized_gain['name'],
                'description': description,
                'debit': 0.00,
                'credit': abs(adj_amount)
            })
        else:
            entries.append({
                'date': adjustment['date'],
                'type': 'FMV Adjustment',
                'asset_name': adjustment['security'],
                'account_code': unrealized_loss['code'],
                'account_name': unrealized_loss['name'],
                'description': description,
                'debit': abs(adj_amount),
                'credit': 0.00
            })
            entries.append({
                'date': adjustment['date'],
                'type': 'FMV Adjustment',
                'asset_name': adjustment['security'],
                'account_code': inv_account['code'],
                'account_name': inv_account['name'],
                'description': description,
                'debit': 0.00,
                'credit': abs(adj_amount)
            })

    return entries


def main():
    args = parse_args()

    # Calculate period_begin if period-type provided
    if args.period_begin:
        period_begin = args.period_begin
    elif args.period_type:
        period_begin = calculate_period_begin(args.through_date, args.period_type)
    else:
        period_begin = None

    # Load registers from configured directories
    asset_register = load_asset_register(args.fixed_assets_dir)
    securities_register = load_securities_register(args.investments_dir)

    all_entries = []

    # Generate entries for all fixed assets
    if period_begin:
        print(f"\nGenerating consolidated journal for {period_begin} through {args.through_date}...")
    else:
        print(f"\nGenerating consolidated journal through {args.through_date}...")

    for asset in asset_register.get('assets', []):
        print(f"  Processing fixed asset: {asset['name']}")

        # Capitalization entry (if purchase was recorded)
        cap_entries = generate_capitalization_entries(asset, args.through_date)
        if cap_entries:
            print(f"    - Including capitalization entry")
            all_entries.extend(cap_entries)

        # Depreciation entries
        entries = generate_depreciation_entries(asset, args.through_date)
        all_entries.extend(entries)

        # Disposal entries (if disposed)
        disposal_entries = generate_disposal_entries_from_asset(asset, args.through_date)
        if disposal_entries:
            print(f"    - Including disposal entries")
            all_entries.extend(disposal_entries)

    # Generate securities transaction entries
    for txn in securities_register.get('transactions', []):
        print(f"  Processing securities: {txn['security']} ({txn['type']})")
        entries = generate_securities_entries(txn, args.through_date)
        all_entries.extend(entries)

    # Generate FMV adjustment entries
    for adj in securities_register.get('fmv_adjustments', []):
        print(f"  Processing FMV adjustment: {adj['security']}")
        entries = generate_fmv_entries(adj, args.through_date)
        all_entries.extend(entries)

    # Sort by date
    all_entries.sort(key=lambda x: x['date'])

    # Filter by period_begin if provided
    if period_begin:
        all_entries = [e for e in all_entries if e['date'] >= period_begin]

    # Save consolidated journal
    os.makedirs(args.output_dir, exist_ok=True)
    csv_path = os.path.join(args.output_dir, f'consolidated_journal_{args.through_date}.csv')
    xlsx_path = os.path.join(args.output_dir, f'consolidated_journal_{args.through_date}.xlsx')

    save_journal_csv(all_entries, csv_path)
    xlsx_ok = save_journal_xlsx(all_entries, xlsx_path, f"Journal through {args.through_date}")

    # Calculate totals
    total_debits = sum(e['debit'] for e in all_entries)
    total_credits = sum(e['credit'] for e in all_entries)

    # Print summary
    print(f"\n{'='*60}")
    print(f"CONSOLIDATED JOURNAL GENERATED")
    print(f"{'='*60}")
    if period_begin:
        print(f"\nPeriod: {period_begin} through {args.through_date}")
    else:
        print(f"\nThrough Date: {args.through_date}")
    print(f"Total Entries: {len(all_entries)}")
    print(f"\n--- Totals ---")
    print(f"Total Debits:  ${total_debits:,.2f}")
    print(f"Total Credits: ${total_credits:,.2f}")
    print(f"Difference:    ${total_debits - total_credits:,.2f}")

    print(f"\n--- Output Files ---")
    print(f"Journal (CSV):  {csv_path}")
    if xlsx_ok:
        print(f"Journal (XLSX): {xlsx_path}")
    print(f"\n{'='*60}")


if __name__ == '__main__':
    main()
