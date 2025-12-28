#!/usr/bin/env python3
"""
run_roll_forward.py - Production GASB 75 Roll-Forward Runner

This script runs a complete roll-forward valuation from start to finish.

Usage:
    python run_roll_forward.py --input prior_year.xlsx --output current_year.xlsx \\
        --prior-date 2024-09-30 --new-date 2025-09-30 \\
        --prior-rate 0.0381 --new-rate 0.0502

Or interactively:
    python run_roll_forward.py --interactive

Author: Actuarial Pipeline Project
Version: 2.1.0 (West Florida Planning Corrections - 2025-12-28)
"""

import argparse
import sys
from datetime import datetime, date
from pathlib import Path

from opeb_rollforward import (
    run_roll_forward,
    verify_roll_forward_output,
    print_roll_forward_summary,
    RollForwardInputs,
)


def parse_date(date_str: str) -> date:
    """Parse date string in YYYY-MM-DD format."""
    return datetime.strptime(date_str, '%Y-%m-%d').date()


def interactive_mode():
    """Run in interactive mode, prompting for all inputs."""
    print("=" * 60)
    print("GASB 75 ROLL-FORWARD VALUATION - Interactive Mode")
    print("=" * 60)
    print()
    
    # Get file paths
    input_path = input("Prior year Excel file path: ").strip()
    if not Path(input_path).exists():
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)
    
    output_path = input("Output file path: ").strip()
    
    # Get dates
    prior_date_str = input("Prior measurement date (YYYY-MM-DD): ").strip()
    prior_date = parse_date(prior_date_str)
    
    new_date_str = input("New measurement date (YYYY-MM-DD): ").strip()
    new_date = parse_date(new_date_str)
    
    # Get discount rates
    prior_rate = float(input("Prior discount rate (e.g., 0.0381 for 3.81%): ").strip())
    new_rate = float(input("New discount rate (e.g., 0.0502 for 5.02%): ").strip())
    
    # Optional parameters
    print()
    print("Optional parameters (press Enter for defaults):")
    
    duration_str = input("Liability duration [10.0]: ").strip()
    duration = float(duration_str) if duration_str else 10.0
    
    trend_duration_str = input("Trend duration [5.0]: ").strip()
    trend_duration = float(trend_duration_str) if trend_duration_str else 5.0
    
    payroll_growth_str = input("Payroll growth rate [0.03]: ").strip()
    payroll_growth = float(payroll_growth_str) if payroll_growth_str else 0.03
    
    benefit_changes = input("Benefit changes description [None]: ").strip()
    if not benefit_changes:
        benefit_changes = "None"
    
    print()
    print("Running roll-forward...")
    print()
    
    # Run the roll-forward
    output, results = run_roll_forward(
        input_path=input_path,
        output_path=output_path,
        prior_measurement_date=prior_date,
        new_measurement_date=new_date,
        prior_discount_rate=prior_rate,
        new_discount_rate=new_rate,
        duration=duration,
        trend_duration=trend_duration,
        payroll_growth_rate=payroll_growth,
        benefit_changes=benefit_changes,
    )
    
    # Create inputs object for summary
    inputs = RollForwardInputs(
        prior_measurement_date=prior_date,
        new_measurement_date=new_date,
        prior_discount_rate=prior_rate,
        new_discount_rate=new_rate,
        boy_tol_old_rate=results.boy_tol_old_rate,
        service_cost=results.service_cost,
        covered_payroll_prior=results.covered_payroll_new / (1 + payroll_growth),
        duration=duration,
        trend_duration=trend_duration,
        payroll_growth_rate=payroll_growth,
        benefit_changes=benefit_changes,
    )
    
    # Print summary
    print_roll_forward_summary(results, inputs)
    
    # Verify output
    print()
    print("Verifying output...")
    verification = verify_roll_forward_output(output)
    
    print()
    print("Verification Results:")
    for check_name, check_result in verification['checks'].items():
        status = "✓ PASS" if check_result['passed'] else "✗ FAIL"
        print(f"  {status}: {check_name}")
        if not check_result['passed']:
            print(f"         Expected: {check_result['expected']}")
            print(f"         Actual: {check_result['actual']}")
    
    print()
    if verification['passed']:
        print("✓ All verification checks passed!")
    else:
        print("✗ Some verification checks failed - please review the output file.")
    
    print()
    print(f"Output saved to: {output}")


def main():
    parser = argparse.ArgumentParser(
        description='Run GASB 75 Roll-Forward Valuation',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode
  python run_roll_forward.py --interactive
  
  # Command line mode
  python run_roll_forward.py \\
      --input "GASB_75_prior_year.xlsx" \\
      --output "GASB_75_current_year.xlsx" \\
      --prior-date 2024-09-30 \\
      --new-date 2025-09-30 \\
      --prior-rate 0.0381 \\
      --new-rate 0.0502
"""
    )
    
    parser.add_argument('--interactive', '-i', action='store_true',
                        help='Run in interactive mode')
    parser.add_argument('--input', type=str, help='Prior year Excel file')
    parser.add_argument('--output', type=str, help='Output Excel file')
    parser.add_argument('--prior-date', type=str, help='Prior measurement date (YYYY-MM-DD)')
    parser.add_argument('--new-date', type=str, help='New measurement date (YYYY-MM-DD)')
    parser.add_argument('--prior-rate', type=float, help='Prior discount rate (e.g., 0.0381)')
    parser.add_argument('--new-rate', type=float, help='New discount rate (e.g., 0.0502)')
    parser.add_argument('--duration', type=float, default=10.0, help='Liability duration (default: 10.0)')
    parser.add_argument('--trend-duration', type=float, default=5.0, help='Trend duration (default: 5.0)')
    parser.add_argument('--payroll-growth', type=float, default=0.03, help='Payroll growth rate (default: 0.03)')
    parser.add_argument('--benefit-changes', type=str, default='None', help='Benefit changes description')
    parser.add_argument('--verify', action='store_true', help='Run verification checks')
    
    args = parser.parse_args()
    
    if args.interactive:
        interactive_mode()
        return
    
    # Check required arguments
    required = ['input', 'output', 'prior_date', 'new_date', 'prior_rate', 'new_rate']
    missing = [arg for arg in required if getattr(args, arg.replace('-', '_')) is None]
    
    if missing:
        print(f"ERROR: Missing required arguments: {', '.join(missing)}")
        print("Use --interactive for guided input or --help for usage.")
        sys.exit(1)
    
    # Parse dates
    prior_date = parse_date(args.prior_date)
    new_date = parse_date(args.new_date)
    
    print("=" * 60)
    print("GASB 75 ROLL-FORWARD VALUATION")
    print("=" * 60)
    print(f"Input:  {args.input}")
    print(f"Output: {args.output}")
    print(f"Period: {prior_date} → {new_date}")
    print(f"Rates:  {args.prior_rate:.2%} → {args.new_rate:.2%}")
    print()
    
    # Run roll-forward
    output, results = run_roll_forward(
        input_path=args.input,
        output_path=args.output,
        prior_measurement_date=prior_date,
        new_measurement_date=new_date,
        prior_discount_rate=args.prior_rate,
        new_discount_rate=args.new_rate,
        duration=args.duration,
        trend_duration=args.trend_duration,
        payroll_growth_rate=args.payroll_growth,
        benefit_changes=args.benefit_changes,
    )
    
    # Create inputs for summary
    inputs = RollForwardInputs(
        prior_measurement_date=prior_date,
        new_measurement_date=new_date,
        prior_discount_rate=args.prior_rate,
        new_discount_rate=args.new_rate,
        boy_tol_old_rate=results.boy_tol_old_rate,
        service_cost=results.service_cost,
        covered_payroll_prior=results.covered_payroll_new / (1 + args.payroll_growth),
        duration=args.duration,
        trend_duration=args.trend_duration,
        payroll_growth_rate=args.payroll_growth,
        benefit_changes=args.benefit_changes,
    )
    
    print_roll_forward_summary(results, inputs)
    
    # Verify if requested
    if args.verify:
        print()
        print("Running verification...")
        verification = verify_roll_forward_output(output)
        
        for check_name, check_result in verification['checks'].items():
            status = "✓" if check_result['passed'] else "✗"
            print(f"  {status} {check_name}: {check_result['actual']}")
        
        if not verification['passed']:
            print()
            print("WARNING: Some verification checks failed!")
            sys.exit(1)
    
    print()
    print(f"✓ Output saved to: {output}")


if __name__ == '__main__':
    main()
