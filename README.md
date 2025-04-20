# PMxl - Project Management Excel Tools

A comprehensive collection of Python scripts and VBA code for generating professional construction and project management Excel templates.

## Table of Contents

-   Overview
-   Project Structure
-   Installation
-   Python Scripts
-   VBA Modules
-   Generated Excel Files
-   How to Contribute

## Overview

PMxl provides ready-to-use Excel templates for construction and project management professionals. The repository includes Python scripts that generate sophisticated Excel files for various project management needs, as well as VBA modules that enhance Excel functionality for specific project management tasks.

## Project Structure

PMxl/

├── py/ # Python scripts for generating Excel templates

├── vba/ # VBA code modules for Excel functionality

├── xl/ # Generated Excel files (output directory)

├── run_scripts.py # Script to execute all Python generators

└── requirements.txt # Python dependencies

## Installation

### Requirements

-   Python 3.8 or higher
-   Required Python packages:
    
    openpyxl>=3.1.0
    
    pandas>=2.0.0
    

### Setup

1.  Clone this repository:
    
    git clone [repository-url]
    
    cd PMxl
    
2.  Install required Python packages:
    
    pip install -r requirements.txt
    

## Python Scripts

All Python scripts are located in the  py  folder. To run all scripts at once:

python run_scripts.py

This will execute all Python generators and save the resulting Excel files to the  xl  folder.

### Available Python Generators

Script

Description

`subcontractor_payment_app_python.py`

Creates a detailed subcontractor payment application form with calculation fields

`python-submittal-excel.py`

Generates a construction submittal tracking workbook with log and templates

`meeting_workbook_python.py`

Produces a meeting management workbook with agenda, minutes, and log sheets

`construction-schedule-python.py`

Creates a Gantt chart-based construction schedule workbook

`construction-rfi-python.py`

Generates an RFI (Request for Information) log and form

`construction-permit-python.py`

Creates a permit tracking log with application forms

`construction-budget-python.py`

Produces a comprehensive budget tracking workbook with forecasting

`construction-bidding-excel-generator.py`

Creates bidding documents including prequalification forms

`construction_safety_program_python.py`

Generates safety program documentation and checklists

`construction_daily_report_generator.py`

Creates daily report templates with dashboard and analytics

To run an individual script:

python py/[script-name].py

## VBA Modules

The VBA modules in the  vba  folder enhance Excel functionality for specific project management tasks. To use these modules:

1.  Open the Excel file where you want to use the VBA code
2.  Press Alt+F11 to open the VBA editor
3.  Right-click on "VBAProject" in the Project Explorer
4.  Select "Import File..."
5.  Navigate to the  vba  folder and select the desired module

### Available VBA Modules

_(List and describe VBA modules here - this would need to be filled in based on the actual VBA files in your project)_

## Generated Excel Files

All generated Excel files are saved to the  xl  folder. Below is a list of the generated files and their descriptions:

-   **Subcontractor Payment Application**  - Comprehensive payment application form with calculation fields
-   **Construction Submittal Workbook**  - Complete submittal tracking system with logs and templates
-   **Meeting Management Workbook**  - Meeting planning and documentation tool
-   **Construction Project Schedule**  - Gantt chart-based project schedule tracker
-   **RFI Log and Form**  - Request for Information tracking system
-   **Construction Permit Log**  - Permit application and tracking system
-   **Construction Budget Workbook**  - Budget planning and tracking with forecasting
-   **Construction Bidding Workbook**  - Complete bidding documentation package
-   **Construction Safety Program**  - Safety documentation and checklist system
-   **Construction Daily Report**  - Daily report system with dashboard and analytics

## How to Contribute

1.  Fork the repository
2.  Create a feature branch (`git checkout -b feature/your-feature`)
3.  Commit your changes (`git commit -m 'Add some feature'`)
4.  Push to the branch (`git push origin feature/your-feature`)
5.  Open a Pull Request