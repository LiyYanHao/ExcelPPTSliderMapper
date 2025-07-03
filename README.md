# Excel-PPT Mapper (ExcelPPTSliderMapper)

A powerful Python tool for automatically mapping Excel data to PowerPoint templates, supporting PTML (PPT Template Markup Language) markup language.

## 🌟 Features

### Core Features
- **Automatic Data Mapping**: Read data from Excel files and automatically replace placeholders in PPT templates
- **Multiple Content Support**: Support text, tables, charts, images, and various PowerPoint content types
- **PTML Markup Language**: Use simple `${variable}` syntax to mark content for replacement
- **Smart File Handling**: Automatic detection and handling of file occupation
- **Process Management**: Automatic PowerPoint process management to avoid conflicts

### Supported Data Types
- **Text Data**: Plain text, percentages, numerical values, etc.
- **Date Data**: Automatic date formatting
- **Table Data**: Dynamic table content replacement
- **Chart Data**: Support for column charts, line charts, combo charts, etc.
- **Image Data**: Dynamic image replacement

### Advanced Features
- **Case Insensitive**: Support case-insensitive key name matching
- **Percentage Smart Processing**: Automatic recognition and formatting of percentage data
- **Combo Chart Support**: Support for column+line combination charts
- **Safe File Saving**: Retry mechanism for safe file saving
- **Detailed Logging**: Complete processing logs

## 📋 System Requirements

### Python Environment
- Python 3.6+
- Windows OS (COM component support required)

### Dependencies
```
win32com.client  # Windows COM component
pandas          # Data processing
openpyxl        # Excel file reading
psutil          # Process management
```

### Microsoft Office
- Microsoft PowerPoint 2016+ 
- Microsoft Excel 2016+

## 🚀 Installation

1. **Clone Project**
```bash
git clone https://github.com/your-repo/ExcelPPTSliderMapper.git
cd ExcelPPTSliderMapper
```

2. **Install Dependencies**
```bash
pip install -r requirements.txt
```

3. **Install pywin32**
```bash
pip install pywin32
```

## 📖 Usage

### Basic Usage

```python
from src.excel_ppt_mapper.excel_ppt_mapper import process_ptml_template, read_excel_template

# File paths
ppt_template = "template.pptx"
excel_data = "data.xlsx" 
output_file = "result.pptx"

# Read Excel data
template_data = read_excel_template(excel_data)

# Process PPT template
success = process_ptml_template(
    ppt_path=ppt_template,
    template_data=template_data,
    output_path=output_file,
    page_numbers=[1, 2, 3]  # Optional: specify pages to process
)

if success:
    print("✅ PPT processing completed")
else:
    print("❌ PPT processing failed")
```

### PTML Markup Syntax

Use the following markup in PowerPoint templates:

```
${variable}  # Direct replacement without any conversion
```

**Examples:**
- `${company_name}` → "Wucang Technology Co., Ltd"
- `${growth_rate}` → "25%"
- `${report_date}` → "2024-01-15"

## 📊 Excel Data Format

### Text Data (Text Sheet)
| Key | Value |
|-----|-------|
| company_name | Wucang Technology Co., Ltd |
| growth_rate | 25% |
| revenue | 10M |

### Date Data (Dates Sheet)
| Key | Value |
|-----|-------|
| report_date | 2024-01-15 |
| end_date | 2024-12-31 |

### Table Data (Tables Sheet)
| Column | Description |
|--------|-------------|
| table_name | Table name |
| header | Header row |
| data | Data rows |

### Chart Data (combo_charts Sheet)
| Chart Name | Category | Series Type | Series Name | Value | Chart Type | Title |
|------------|----------|-------------|-------------|-------|------------|-------|
| sales_chart | Q1 | column | Sales | 100 | combo | Quarterly Sales Analysis |
| sales_chart | Q1 | line | Growth Rate | 15 | combo | Quarterly Sales Analysis |

### Image Data (Images Sheet)
| Key | Value |
|-----|-------|
| company_logo | ./images/logo.png |
| product_image | ./images/product.jpg |

## 🔧 API Documentation

### Main Functions

#### `process_ptml_template(ppt_path, template_data, output_path, page_numbers)`
Main function for processing PPT templates

**Parameters:**
- `ppt_path` (str): PPT template file path
- `template_data` (Dict): Template data dictionary
- `output_path` (str, optional): Output file path
- `page_numbers` (List[int], optional): List of page numbers to process

**Returns:**
- `bool`: Whether processing was successful

#### `read_excel_template(excel_path)`
Read template data from Excel file

**Parameters:**
- `excel_path` (str): Excel file path

**Returns:**
- `Dict[str, Any]`: Template data dictionary

### Helper Functions

#### `check_file_in_use(file_path)`
Check if a file is being used

#### `close_powerpoint_processes()`
Close all PowerPoint processes

#### `safe_save_presentation(pres, output_path, max_retries)`
Safely save presentation with retry mechanism

## 💡 Usage Examples

### Example 1: Process Complete PPT
```python
# Process all pages
template_data = read_excel_template("data.xlsx")
process_ptml_template("template.pptx", template_data, "output.pptx")
```

### Example 2: Process Specific Pages
```python
# Only process pages 1-3
template_data = read_excel_template("data.xlsx")
process_ptml_template(
    "template.pptx", 
    template_data, 
    "output.pptx", 
    page_numbers=[1, 2, 3]
)
```

### Example 3: Text Replacement
**In PPT Template:**
```
Company: ${company_name}
Growth Rate: ${growth_rate}
```

**Excel Data:**
| Key | Value |
|-----|-------|
| company_name | Wucang Tech |
| growth_rate | 25% |

**Result:**
```
Company: Wucang Tech
Growth Rate: 25%
```

## ⚠️ Notes

### File Operations
1. **File Permissions**: Ensure sufficient read/write permissions
2. **File Occupation**: Program automatically handles file occupation issues
3. **Backup Important Files**: Recommend backing up original files before processing

### PowerPoint Related
1. **Version Compatibility**: Recommend Office 2016 or above
2. **Macro Security**: May need to adjust PowerPoint macro security settings
3. **Process Cleanup**: Program automatically cleans up PowerPoint processes

### Data Format
1. **Encoding Issues**: Recommend using UTF-8 encoding for Excel files
2. **Data Types**: Ensure consistency in number, text, and date formats
3. **Special Characters**: Avoid special characters in key names

### Performance Optimization
1. **Large File Processing**: Large PPT files may require longer processing time
2. **Memory Usage**: Complex charts may use more memory
3. **Concurrency Limitation**: Not recommended to run multiple instances simultaneously

## 🐛 Troubleshooting

### Common Issues

**Q: "File in use" message**
A: Program will automatically close PowerPoint processes; if issues persist, manually close all Office programs

**Q: Markers not replaced**
A: Check if Excel key names exactly match PPT markers (case-insensitive matching supported)

**Q: Chart data update failed**
A: Ensure correct chart data format in Excel, especially for numerical values

**Q: Program running slowly**
A: May be due to large PPT files or complex charts, please be patient

### Debug Mode
Program provides detailed log output, check console output for diagnostics.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

## 🤝 Contributing

Issues and Pull Requests are welcome to help improve this project!

## 📞 Support

If you encounter any problems or have feature suggestions, please:
1. Check the troubleshooting section of this document
2. Search existing Issues
3. Create a new Issue

---

**Note**: This tool requires Windows environment and depends on Microsoft Office suite. 