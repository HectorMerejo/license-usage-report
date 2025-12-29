# Azure License Usage Report

Automated script that analyzes Microsoft license usage in Azure AD (Entra ID) and identifies potential cost savings from unused licenses.

## Features

- Retrieves all license subscriptions (SKUs) from Azure tenant
- Analyzes license utilization across organization
- Identifies underutilized licenses (potential waste)
- Generates formatted Excel report with:
  - Summary dashboard with color-coded utilization
  - Detailed user assignment breakdown
- Provides cost-saving recommendations

## Output

- **Excel file** with two sheets:
  - **License Summary**: Overview with utilization percentages (color-coded: 🟢 Green = good utilization, 🟡 Yellow = moderate, 🔴 Red = wasteful)
  - **User Assignments**: Detailed list of which users have which licenses

## Requirements

- Python 3.8+
- Azure AD app registration with API permissions:
  - `User.Read.All`
  - `Organization.Read.All`
- Microsoft Graph API access

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/YOUR-USERNAME/azure-license-usage-report.git
cd azure-license-usage-report
```

### 2. Create virtual environment

```bash
# Mac/Linux
python3 -m venv venv
source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure Azure credentials

Create a `.env` file in the project root:

```
CLIENT_ID=your-application-id
CLIENT_SECRET=your-client-secret
TENANT_ID=your-tenant-id
```

To get these credentials:

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to Microsoft Entra ID → App registrations
3. Create a new app registration
4. Grant API permissions: `User.Read.All`, `Organization.Read.All`
5. Create a client secret
6. Copy the Application (client) ID, Directory (tenant) ID, and secret value

### 5. Run the script

```bash
# Mac/Linux
python3 license_report.py

# Windows
python license_report.py
```

## Business Value

Helps organizations optimize software licensing costs by:

- **Identifying unused licenses** - Find licenses that are purchased but not assigned
- **Analyzing utilization rates** - See which licenses are underutilized
- **Supporting budget decisions** - Provide data for procurement and renewals
- **Maintaining compliance** - Create audit trails of license assignments
- **Estimating cost savings** - Highlight potential savings from reducing unused licenses

### Example Cost Impact

If your organization has:

- 50 Microsoft 365 E3 licenses @ $36/month = $1,800/month
- Only 35 licenses actually assigned = 15 unused licenses
- **Potential savings**: 15 × $36 × 12 months = **$6,480/year**

## Report Interpretation

### Color Coding

- 🟢 **Green (80%+ utilization)**: Optimal - licenses are being used effectively
- 🟡 **Yellow (50-80% utilization)**: Moderate - some waste, monitor for optimization
- 🔴 **Red (<50% utilization)**: High waste - strong candidate for cost reduction

### Sample Output

```
======================================================================
LICENSE USAGE SUMMARY
======================================================================
Total Licenses Purchased:  150
Total Assigned:            108
Total Available (Unused):  42
Overall Utilization:       72.0%

⚠️  LOW UTILIZATION LICENSES (Potential Cost Savings):
----------------------------------------------------------------------
Office 365 E3                            | 15 unused (70% utilization)
Power BI Pro                             | 10 unused (50% utilization)
```

## Technologies Used

- **Python 3** - Core programming language
- **Microsoft Graph API** - Azure data retrieval
- **MSAL** (Microsoft Authentication Library) - Secure authentication
- **openpyxl** - Excel file generation with formatting
- **Azure Active Directory** (Entra ID) - Identity management platform

## Security Notes

- Never commit your `.env` file to version control
- Store credentials securely
- Use service accounts with minimum required permissions
- Rotate client secrets regularly (recommend 90-day rotation)
- Review API permissions match organizational security policies

## Troubleshooting

### "Authentication failed"

- Verify CLIENT_ID, CLIENT_SECRET, and TENANT_ID in `.env` file
- Ensure no extra spaces in credentials
- Check that client secret hasn't expired

### "Insufficient privileges"

- Go to Azure Portal → App registrations → API permissions
- Ensure `User.Read.All` and `Organization.Read.All` are granted
- Click "Grant admin consent for [Your Organization]"

### "No licenses found"

- This is normal for personal Azure accounts (no paid subscriptions)
- Script will work correctly in organizational Azure tenants
- Test in a work/production environment to see license data

## Use Cases

- **IT Administrators**: Monthly license audits and optimization
- **Finance Teams**: Budget planning and cost reduction initiatives
- **Compliance Officers**: License usage documentation for audits
- **Procurement**: Data-driven decisions on license renewals

## Future Enhancements

- Email report delivery
- Scheduled automation (weekly/monthly)
- Historical trending (track license usage over time)
- Cost calculations based on license pricing
- Recommendations engine for license optimization
- Integration with procurement systems

## Author

Built as part of Azure administration automation toolkit.

## License

MIT License - free to use and modify for your organization's needs.
