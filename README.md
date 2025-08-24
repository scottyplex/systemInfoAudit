# Windows System Audit Script

This PowerShell script automates the process of collecting key system information from a Windows host and compiling it into a single, multi-tab Excel file. It is designed to save time during server auditing, troubleshooting, and asset management.

---

### Features

* **Comprehensive Data Collection:** Gathers essential system details, including hardware configuration, network settings, disk space, and a list of all installed programs.
* **Single-File Report:** Combines all collected data from multiple sources into a single, easy-to-read Excel workbook with dedicated tabs for each data category.
* **Automated Cleanup:** Automatically deletes all temporary CSV files after the final Excel report is generated, keeping your desktop clean.
* **User-Friendly Output:** Opens the final Excel report automatically upon completion, so the information is ready for immediate review.

---

### Prerequisites

* A Windows host with **PowerShell 5.1 or newer**.
* **Microsoft Excel** must be installed on the host to enable the script to merge CSV files into a multi-tab workbook.

---

### Usage

Simply run the script. It does not require any parameters.

```
.\Get-SystemInfo.ps1
```

The script will create an `audit` folder on your desktop and place the final Excel file, named `ServerAudit.xlsx`, inside it.

---

### Output

The final Excel file will contain the following tabs:

* **SystemInfo**: Basic system details like the computer name, domain, manufacturer, and hardware configuration.
* **Network**: A full network configuration report from `ipconfig /all`.
* **HDDConfig**: Information on all logical disks and their file systems, size, and free space.
* **InstalledPrograms**: A comprehensive list of all installed software, including the display name, version, and publisher.
