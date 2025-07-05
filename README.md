# Excel to vCard Converter

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A powerful and flexible Python tool for converting Excel files to vCard format (.vcf). Perfect for importing contact data into mobile phones, email clients, and contact management systems.

## ✨ Features

- **Two Conversion Modes**: Full-featured and simple conversion options
- **Flexible Input**: Support for various Excel formats (.xlsx, .xls)
- **Multiple Fields**: Handle name, phone, email, organization, title, and more
- **Error Handling**: Robust error handling and validation
- **Command Line Interface**: Easy-to-use CLI with customizable options
- **UTF-8 Support**: Proper encoding for international characters

## 📋 Requirements

- Python 3.7 or higher
- Required packages (see `requirements.txt`):
  - `pandas` >= 1.5.0
  - `xlrd` >= 2.0.0
  - `openpyxl` >= 3.0.0

## 🚀 Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/Excel-to-Vcard-converter.git
   cd Excel-to-Vcard-converter
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## 📖 Usage

### Basic Usage

**Full-featured converter** (recommended):
```bash
python excel_to_vcard.py Contacts.xlsx
```

**Simple converter** (basic fields only):
```bash
python simple_excel_to_vcard.py Contacts.xlsx
```

### Advanced Usage

**Specify custom sheet and output file**:
```bash
python excel_to_vcard.py data.xlsx --sheet "Contacts" --output "my_contacts.vcf"
```

**Simple converter with custom options**:
```bash
python simple_excel_to_vcard.py data.xlsx --sheet "Employees" --output "employees.vcf"
```

### Command Line Options

| Option | Short | Description | Default |
|--------|-------|-------------|---------|
| `excel_file` | - | Path to the Excel file | Required |
| `--sheet` | `-s` | Sheet name to process | "Workers" |
| `--output` | `-o` | Output vCard file name | "Exported.vcf" |

## 📊 Excel File Format

### Required Columns

The converter expects specific column names in your Excel file:

| Column Name | Description | Required |
|-------------|-------------|----------|
| `Phone` | Contact phone number | ✅ Yes |
| `Name` | First name | ✅ Yes |
| `Surname` | Last name | ✅ Yes |

### Optional Columns

| Column Name | Description | Converter |
|-------------|-------------|-----------|
| `Mail` | Email address | Both |
| `MiddleName` | Middle name | Full only |
| `Prefix` | Name prefix (Mr., Dr., etc.) | Full only |
| `Suffix` | Name suffix (Jr., Sr., etc.) | Full only |
| `Organization` | Company/organization | Full only |
| `Title` | Job title | Full only |

### Example Excel Structure

| Name | Surname | Phone | Mail | Organization | Title |
|------|---------|-------|------|--------------|-------|
| John | Doe | +1234567890 | john@example.com | Acme Corp | Manager |
| Jane | Smith | +0987654321 | jane@example.com | Tech Inc | Developer |

## 🔧 Converter Comparison

### Full Converter (`excel_to_vcard.py`)
- **Features**: All contact fields supported
- **Use case**: Complete contact information
- **Output**: Rich vCard with all available data

### Simple Converter (`simple_excel_to_vcard.py`)
- **Features**: Basic fields only (Name, Surname, Phone, Email)
- **Use case**: Quick conversion with minimal data
- **Output**: Streamlined vCard format

## 📁 Output Format

The converter generates a standard vCard (.vcf) file:

```vcf
BEGIN:VCARD
VERSION:2.1
N:Doe;John; ; ;
FN:John Doe
TEL;CELL:+1234567890
EMAIL;HOME:john@example.com
ORG:Acme Corp
TITLE:Manager
END:VCARD
```

## 🛠️ Development

### Project Structure
```
Excel-to-Vcard-converter/
├── excel_to_vcard.py          # Full-featured converter
├── simple_excel_to_vcard.py   # Simple converter
├── requirements.txt           # Python dependencies
├── LICENSE                    # MIT License
├── README.md                  # This file
└── Contacts.xlsx             # Example Excel file
```

### Running Tests
```bash
# Test with example file
python excel_to_vcard.py Contacts.xlsx
python simple_excel_to_vcard.py Contacts.xlsx
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'feat: add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Commit Convention

This project follows [Conventional Commits](https://www.conventionalcommits.org/):

- `feat:` New features
- `fix:` Bug fixes
- `docs:` Documentation changes
- `style:` Code style changes
- `refactor:` Code refactoring
- `test:` Test additions or changes
- `chore:` Maintenance tasks

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🐛 Troubleshooting

### Common Issues

**"File not found" error**:
- Ensure the Excel file exists in the specified path
- Check file permissions

**"Sheet not found" error**:
- Verify the sheet name matches exactly (case-sensitive)
- Use `--sheet` option to specify the correct sheet

**"No contacts converted"**:
- Check that the "Phone" column exists and contains data
- Ensure required columns are present

**Encoding issues**:
- The converter uses UTF-8 encoding by default
- Ensure your Excel file doesn't contain invalid characters

### Getting Help

If you encounter issues:
1. Check the error messages for specific details
2. Verify your Excel file format matches the requirements
3. Open an issue with your Excel file structure and error details

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/Excel-to-Vcard-converter/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/Excel-to-Vcard-converter/discussions)

---

**Made with ❤️ for easy contact management**



