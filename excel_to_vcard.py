#!/usr/bin/env python3
"""
Excel to vCard Converter

A Python script to convert Excel files to vCard format (.vcf).
Supports multiple contact fields and provides flexible configuration.
"""

import pandas as pd
import os
import sys
from typing import Dict, List, Optional
import argparse


class ExcelToVcardConverter:
    """Main converter class for Excel to vCard conversion."""
    
    def __init__(self, excel_file: str, sheet_name: str = "Workers", output_file: str = "Exported.vcf"):
        """
        Initialize the converter.
        
        Args:
            excel_file (str): Path to the Excel file
            sheet_name (str): Name of the sheet to process
            output_file (str): Output vCard file name
        """
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.output_file = output_file
        self.data = None
        
    def load_excel_data(self) -> bool:
        """
        Load data from Excel file.
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            if not os.path.exists(self.excel_file):
                print(f"Error: File '{self.excel_file}' not found.")
                return False
                
            excel_file = pd.ExcelFile(self.excel_file)
            
            if self.sheet_name not in excel_file.sheet_names:
                print(f"Error: Sheet '{self.sheet_name}' not found in Excel file.")
                print(f"Available sheets: {excel_file.sheet_names}")
                return False
                
            self.data = excel_file.parse(self.sheet_name)
            print(f"Successfully loaded {len(self.data)} contacts from '{self.sheet_name}' sheet.")
            return True
            
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False
    
    def clean_value(self, value) -> str:
        """
        Clean and validate a cell value.
        
        Args:
            value: Raw cell value
            
        Returns:
            str: Cleaned value or empty string
        """
        if pd.isna(value) or str(value).lower() == 'nan':
            return ""
        return str(value).strip()
    
    def create_vcard_entry(self, row: pd.Series) -> str:
        """
        Create a vCard entry from a data row.
        
        Args:
            row (pd.Series): Row data from Excel
            
        Returns:
            str: vCard formatted string
        """
        # Extract and clean values
        phone = self.clean_value(row.get("Phone", ""))
        if not phone:
            return ""  # Skip entries without phone
            
        name = self.clean_value(row.get("Name", ""))
        surname = self.clean_value(row.get("Surname", ""))
        middle_name = self.clean_value(row.get("MiddleName", ""))
        prefix = self.clean_value(row.get("Prefix", ""))
        suffix = self.clean_value(row.get("Suffix", ""))
        email = self.clean_value(row.get("Mail", ""))
        organization = self.clean_value(row.get("Organization", ""))
        title = self.clean_value(row.get("Title", ""))
        
        # Build vCard
        vcard = "BEGIN:VCARD\nVERSION:2.1\n"
        
        # Name field (N:LastName;FirstName;MiddleName;Prefix;Suffix)
        vcard += f"N:{surname};{name};{middle_name};{prefix};{suffix}\n"
        
        # Full name field (FN:DisplayName)
        display_name_parts = [prefix, name, middle_name, surname, suffix]
        display_name = " ".join(part for part in display_name_parts if part)
        vcard += f"FN:{display_name}\n"
        
        # Phone field
        vcard += f"TEL;CELL:{phone}\n"
        
        # Email field
        if email:
            vcard += f"EMAIL;HOME:{email}\n"
        
        # Organization field
        if organization:
            vcard += f"ORG:{organization}\n"
        
        # Title field
        if title:
            vcard += f"TITLE:{title}\n"
        
        vcard += "END:VCARD\n"
        return vcard
    
    def convert(self) -> bool:
        """
        Convert Excel data to vCard format.
        
        Returns:
            bool: True if successful, False otherwise
        """
        if self.data is None:
            print("Error: No data loaded. Call load_excel_data() first.")
            return False
        
        try:
            vcard_content = ""
            processed_count = 0
            
            for index, row in self.data.iterrows():
                vcard_entry = self.create_vcard_entry(row)
                if vcard_entry:
                    vcard_content += vcard_entry
                    processed_count += 1
            
            # Write to file
            with open(self.output_file, "w", encoding="utf-8") as f:
                f.write(vcard_content)
            
            print(f"Successfully converted {processed_count} contacts to '{self.output_file}'")
            return True
            
        except Exception as e:
            print(f"Error during conversion: {e}")
            return False
    
    def run(self) -> bool:
        """
        Run the complete conversion process.
        
        Returns:
            bool: True if successful, False otherwise
        """
        if not self.load_excel_data():
            return False
        
        return self.convert()


def main():
    """Main function to run the converter."""
    parser = argparse.ArgumentParser(
        description="Convert Excel file to vCard format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python excel_to_vcard.py Contacts.xlsx
  python excel_to_vcard.py data.xlsx --sheet "Contacts" --output "my_contacts.vcf"
        """
    )
    
    parser.add_argument("excel_file", help="Path to the Excel file")
    parser.add_argument("--sheet", "-s", default="Workers", help="Sheet name to process (default: Workers)")
    parser.add_argument("--output", "-o", default="Exported.vcf", help="Output vCard file name (default: Exported.vcf)")
    
    args = parser.parse_args()
    
    # Create converter and run
    converter = ExcelToVcardConverter(
        excel_file=args.excel_file,
        sheet_name=args.sheet,
        output_file=args.output
    )
    
    success = converter.run()
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main() 