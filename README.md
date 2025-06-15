# Warehouse Operations Management System

![Application Screenshot](screenshot.png) <!-- Add actual screenshot file later -->

A C# desktop application for managing warehouse operations including data import, supplier assignment, label generation, and reporting.

## Features

- **Excel Data Import**: Import subscriber data from HFMR526 formatted Excel files
- **Automatic Supplier Assignment**: Match imported data with suppliers from PAGEANT ROUTING GUIDE
- **Label Generation**: Produce carrier-compliant labels in INTL Template format
- **Summary Reporting**: Generate reports showing total items received by supplier
- **Secure Data Storage**: Maintain historical records of all transactions

## Prerequisites

- .NET Framework 4.7.2 or later
- Microsoft Excel or Excel Data Reader libraries
- (If using RDLC) Microsoft Report Viewer
- (If using Crystal Reports) Crystal Reports runtime

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/warehouse-ops-system.git
