"""
Excel Comparison Tool - Main Entry Point
"""

from src.core import ComparisonEngine, ComparisonConfig, AlignmentMethod
from src.reports.report_generator import generate_comparison_report
import pandas as pd
from datetime import datetime
import sys
import os
import platform


def main():
    """Main execution function"""
    
    print("=" * 80)
    print("Excel Comparison Tool v1.0")
    print("=" * 80)
    
    # 1. Load your Excel files
    print("\n📂 Loading files...")
    
    try:
        filea = r"C:\VP\tc01_filea.xlsx"
        fileb = r"C:\VP\tc01_fileb.xlsx"
        
        df_a = pd.read_excel(filea)
        df_b = pd.read_excel(fileb)
        
        print(f"✅ File A loaded: {len(df_a)} rows")
        print(f"✅ File B loaded: {len(df_b)} rows")
        
    except FileNotFoundError as e:
        print(f"❌ Error: File not found - {e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error loading files: {e}")
        sys.exit(1)
    
    # 2. Configure comparison
    print("\n⚙️ Configuring comparison...")
    
    config = ComparisonConfig(
        key_columns=['Pol #','Insured','Eff Date'],  # Your composite key
        alignment_method=AlignmentMethod.POSITION,
        case_sensitive=False,
        trim_whitespace=True
    )
    
    print(f"   Key columns: {', '.join(config.key_columns)}")
    print(f"   Alignment method: {config.alignment_method.value}")
    
    # 3. Run comparison
    print("\n🔍 Comparing files...")
    
    try:
        engine = ComparisonEngine(config)
        result = engine.compare(df_a, df_b)
        
        print("✅ Comparison complete!")
        
    except KeyError as e:
        print(f"❌ Error: {e}")
        print("   Make sure the key columns exist in both files.")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Comparison error: {e}")
        sys.exit(1)
    
    # 4. Display summary
    print("\n" + "=" * 80)
    print("📊 COMPARISON SUMMARY")
    print("=" * 80)
    
    summary = result.summary
    print(f"\n🔑 Keys:")
    print(f"   Total unique keys in File A: {summary['total_unique_keys_a']}")
    print(f"   Total unique keys in File B: {summary['total_unique_keys_b']}")
    print(f"   Keys in common: {summary['keys_in_common']}")
    print(f"   Keys only in A: {summary['keys_only_in_a']}")
    print(f"   Keys only in B: {summary['keys_only_in_b']}")
    
    print(f"\n📝 Rows:")
    print(f"   Total rows compared: {summary['total_rows_compared']}")
    print(f"   ✅ Matching rows: {summary['match_count']}")
    print(f"   🟡 Modified rows: {summary['modified_count']}")
    print(f"   🟢 Added rows: {summary['added_row_count']}")
    print(f"   🔴 Removed rows: {summary['removed_row_count']}")
    print(f"   🔵 Rows in new keys: {summary['new_key_count']}")
    print(f"   🟠 Rows in removed keys: {summary['removed_key_count']}")
    
    # 5. Generate Excel report
    print("\n📄 Generating Excel report...")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"comparison_report_{timestamp}.xlsx"
    
    try:
        generate_comparison_report(
            output_path=output_file,
            summary=result.summary,
            aligned_data=result.aligned_data,
            metadata=result.comparison_metadata,
            file_a_path=filea,
            file_b_path=fileb
        )
        
        # Get absolute path
        from pathlib import Path
        absolute_path = Path(output_file).resolve()
        
        print(f"\n✅ SUCCESS! Report saved to:")
        print(f"   {absolute_path}")
        
        # Optional: Auto-open the report
        try:
            if platform.system() == 'Windows':
                os.startfile(absolute_path)
                print(f"\n📂 Opening report in Excel...")
            elif platform.system() == 'Darwin':  # macOS
                os.system(f'open "{absolute_path}"')
            else:  # Linux
                os.system(f'xdg-open "{absolute_path}"')
        except Exception as e:
            print(f"\n   (Could not auto-open file: {e})")
            print(f"   Please open manually: {absolute_path}")
        
    except Exception as e:
        print(f"❌ Error generating report: {e}")
        sys.exit(1)
    
    # 6. Show sample of differences
    if summary['modified_count'] > 0 or summary['added_row_count'] > 0 or summary['removed_row_count'] > 0:
        print("\n" + "=" * 80)
        print("📋 SAMPLE DIFFERENCES (first 5)")
        print("=" * 80)
        
        diff_rows = result.aligned_data[
            result.aligned_data['status'].isin(['MODIFIED', 'ADDED_ROW', 'REMOVED_ROW'])
        ].head(5)
        
        for idx, row in diff_rows.iterrows():
            key_cols = [col for col in row.index if col.startswith('key_')]
            key_str = " | ".join([f"{col.replace('key_', '')}: {row[col]}" for col in key_cols])
            
            print(f"\n   {row['status']}")
            print(f"   Key: {key_str}")
            
            if row['status'] == 'MODIFIED' and 'changed_cells' in row:
                print(f"   Changed: {row['changed_cells']}")
    
    print("\n" + "=" * 80)
    print("🎉 Comparison complete! Open the Excel file to see detailed results.")
    print("=" * 80)


if __name__ == "__main__":
    main()