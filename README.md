# Shipping Data Processing System

A Google Apps Script-based system for automating shipping data processing, cleaning, and reporting. The system handles CSV and PDF attachments from Gmail, processes them, and generates automated reports.

## Features

- **Automated Data Ingestion**
  - Monitors Gmail for CSV and PDF attachments
  - Processes attachments automatically
  - Maintains processing history with Gmail labels

- **Data Processing**
  - Cleans and standardizes incoming data
  - Removes duplicates
  - Maps vendor tiers
  - Handles special shipping requirements

- **Reporting**
  - Generates customized reports by sender
  - Sends automated email reports
  - Maintains email sending logs
  - Supports PDF text extraction

## System Architecture

```
410Shipping/
├── core/                    # Core system components
│   ├── config.js           # System configuration
│   ├── utils.js            # Shared utilities
│   ├── gmail.js            # Gmail operations
│   └── sheets.js           # Google Sheets operations
├── services/               # Business logic services
│   ├── ingestion.js        # Data ingestion service
│   ├── cleaning.js         # Data cleaning service
│   └── reporting.js        # Reporting service
├── templates/              # Configuration templates
│   ├── config_template.csv # System configuration template
│   └── tier_mapping_template.csv # Tier mapping template
└── main.js                 # Main entry point
```

## Setup Instructions

1. **Google Apps Script Setup**
   - Create a new Google Apps Script project
   - Copy all files maintaining the directory structure
   - Enable required Google services:
     - Gmail API
     - Google Drive API
     - Google Sheets API

2. **Configuration**
   - Use `templates/config_template.csv` to configure the system
   - Update the following settings:
     - Spreadsheet ID
     - Email addresses
     - Gmail search parameters
     - Sheet names

3. **Tier Mapping**
   - Use `templates/tier_mapping_template.csv` to define shipping tiers
   - Customize tier parameters:
     - Priority levels
     - Processing times
     - Weight/volume limits
     - Special handling requirements

4. **Google Sheet Setup**
   - Create a new Google Sheet
   - Create the following sheets:
     - Raw Data
     - Clean Data
     - Cols to Send
     - Mail Log
     - Raw Import

## Usage

### Running the Pipeline

```javascript
// Run the complete pipeline
SHIPPING_SYSTEM.runShippingPipeline();

// Run individual steps
SHIPPING_SYSTEM.runIngestion();
SHIPPING_SYSTEM.runCleaning();
SHIPPING_SYSTEM.runReporting();
```

### Maintenance Functions

```javascript
// Clean up processed labels
SHIPPING_SYSTEM.cleanupLabels();

// Remove duplicates
SHIPPING_SYSTEM.removeDuplicates();

// Test configuration
SHIPPING_SYSTEM.testConfiguration();
```

## Configuration

### System Configuration (config_template.csv)

| Category | Setting | Description |
|----------|---------|-------------|
| Spreadsheet | ID | Google Sheet ID |
| Spreadsheet | Sheet Names | Names of required sheets |
| Gmail | Search Query | Email search parameters |
| Gmail | Date Range | Days to look back |
| Data | Tier Column | Column name for tier info |
| Email | Recipient | Report recipient email |

### Tier Mapping (tier_mapping_template.csv)

| Column | Description |
|--------|-------------|
| Tier Code | Letter code (A-D) |
| Tier Name | Full tier name |
| Priority Level | Numeric priority |
| Processing Time | Days to process |
| Max Weight | Weight limit (kg) |
| Max Volume | Volume limit (m³) |
| Special Handling | Handling requirements |

## Error Handling

The system includes comprehensive error handling:
- Logs all operations
- Maintains processing status
- Handles attachment errors
- Manages duplicate data
- Tracks email sending

## Maintenance

### Regular Tasks
- Monitor email logs
- Check processing status
- Review error logs
- Update tier mappings
- Clean up old labels

### Troubleshooting
1. Check Gmail labels
2. Verify spreadsheet access
3. Review processing logs
4. Test configuration
5. Validate tier mappings

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support, please:
1. Check the documentation
2. Review error logs
3. Test configuration
4. Contact system administrator

## Version History

- v1.0.0 (2024-03-20)
  - Initial release
  - Basic pipeline implementation
  - CSV and PDF processing
  - Automated reporting 