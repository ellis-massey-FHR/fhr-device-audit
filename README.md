# FHR Computer Refresh List last 60 Days


This script generates a detailed Excel report of laptop or Surface device requests from ServiceNow over the past 60 days, enriched with user workstation data. It includes a summary pie chart of requests by location and filters only valid, corporate-managed Windows workstations. This helps easliy identify users with multiple PC's.

## 📋 Features

- Queries ServiceNow request items for laptops and Surface devices
- Filters:
  - Non-new hire requests
  - Devices requested or closed in the last 60 days
  - Workstations running Windows OS only
  - Device trust types: AzureAD and HybridAD
- Pulls user workstation data from the CMDB (`cmdb_ci_computer`)
- Ignores mobile and retired devices
- Outputs an Excel file with:
  - Cleaned and formatted data
  - Auto-adjusted columns
  - Title header
  - Embedded pie chart of requests grouped by location

---

## 🛠️ Requirements

- Python 3.8+
- Access to ServiceNow API with basic authentication
- `.env` file with your credentials (see below)
- Required Python libraries:
  ```bash
  pip install requests python-dotenv pandas matplotlib openpyxl
