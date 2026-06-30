# OCI Compute Inventory Report Generator

This Python script automates the collection of Oracle Cloud Infrastructure (OCI) Compute instance details across all accessible compartments within a tenancy and exports the information into a structured Excel report.

The generated report is organized into separate worksheets based on the **Environment** tag (**PROD**, **DEV**, **UAT**) and includes a **Summary** sheet with overall instance statistics.

## Information Collected

- Instance Name
- Instance OCID
- Private IP
- Public IP
- Compartment
- Subnet
- OCPU Count
- Memory (GB)
- Shape
- Operating System
- Lifecycle State
- Environment Tag
- Application Tag
- Creation Time

## Report Features

- Summary dashboard
- Environment-wise worksheets (PROD, DEV, UAT)
- Bold headers
- Auto-adjusted column widths
- Auto filters
- Frozen header row
- Color-coded lifecycle states

## Installation

```bash
pip install -r requirements.txt
```

## Run

```bash
python OCI_Compute_inventory.py
```

After successful execution, a timestamped Excel report will be generated in the project directory.

Example:

```text
OCI_Compute_Report_2026-06-30_10-45.xlsx
```
