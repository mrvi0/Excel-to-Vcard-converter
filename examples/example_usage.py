#!/usr/bin/env python3
"""
Example usage of Excel to vCard converter

This file demonstrates how to use the converter programmatically.
"""

import sys
import os

# Add parent directory to path to import the converter
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_to_vcard import ExcelToVcardConverter
from simple_excel_to_vcard import SimpleExcelToVcardConverter


def example_full_converter():
    """Example of using the full-featured converter."""
    print("=== Full Converter Example ===")
    
    # Create converter instance
    converter = ExcelToVcardConverter(
        excel_file="Contacts.xlsx",
        sheet_name="Workers",
        output_file="full_contacts.vcf"
    )
    
    # Run conversion
    success = converter.run()
    
    if success:
        print("✅ Full conversion completed successfully!")
    else:
        print("❌ Full conversion failed!")


def example_simple_converter():
    """Example of using the simple converter."""
    print("\n=== Simple Converter Example ===")
    
    # Create converter instance
    converter = SimpleExcelToVcardConverter(
        excel_file="Contacts.xlsx",
        sheet_name="Workers",
        output_file="simple_contacts.vcf"
    )
    
    # Run conversion
    success = converter.run()
    
    if success:
        print("✅ Simple conversion completed successfully!")
    else:
        print("❌ Simple conversion failed!")


def example_custom_converter():
    """Example of using the converter with custom settings."""
    print("\n=== Custom Converter Example ===")
    
    # Create converter with custom settings
    converter = ExcelToVcardConverter(
        excel_file="Contacts.xlsx",
        sheet_name="Workers",
        output_file="custom_contacts.vcf"
    )
    
    # Load data manually
    if converter.load_excel_data():
        if converter.data is not None:
            print(f"📊 Loaded {len(converter.data)} contacts")
            
            # Show available columns
            print(f"📋 Available columns: {list(converter.data.columns)}")
        else:
            print("❌ No data loaded!")
        
        # Convert
        if converter.convert():
            print("✅ Custom conversion completed!")
        else:
            print("❌ Custom conversion failed!")
    else:
        print("❌ Failed to load data!")


if __name__ == "__main__":
    print("Excel to vCard Converter - Examples")
    print("=" * 40)
    
    # Check if example file exists
    if not os.path.exists("Contacts.xlsx"):
        print("⚠️  Warning: Contacts.xlsx not found in current directory")
        print("   Please ensure you have an Excel file to convert")
        sys.exit(1)
    
    # Run examples
    example_full_converter()
    example_simple_converter()
    example_custom_converter()
    
    print("\n🎉 All examples completed!")
    print("Check the generated .vcf files in the current directory.") 