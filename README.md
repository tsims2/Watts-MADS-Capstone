# University of Michigan MADS Capstone - Watts
### Automation and Data Analysis with Advance GUI

Hello! Please read this file before using the Watts Demo. 
This documents will walk you through the demo and the outputs it produces. 

## Overview

Watts is a comprehensive energy management system designed for managing and analyzing energy data, particularly for Connected Solutions programs. 
This application provides tools for data visualization, performance calculations, and report generation for energy assets and events.
This demo has been reduced in scoped and refined for the purpose of this capstone.

## Features

- **Data Management**: Import, view, and manage energy data from various sources.
- **Performance Calculations**: Calculate and analyze performance metrics for energy assets.
- **Event Management**: Track and manage energy events, including daily and targeted dispatch events.
- **Visualization Tools**: 
  - Heatmap for performance cross-analysis
  - Geo-analysis for spatial distribution of assets
  - Line graphs for baseline and event performance
- **Report Generation**: Create detailed PDF reports for individual assets and customers.
- **User Interface**: 
  - Tree view for data display and selection
  - Data viewer for visual analysis
  - Filterable data table

## Tech Stack

- **Python**: Primary programming language
- **PyQt5**: GUI framework
- **Pandas & NumPy**: Data manipulation and analysis
- **Matplotlib & Seaborn**: Data visualization
- **Folium**: Interactive map creation
- **ReportLab & PyPDF2**: PDF generation and manipulation
- **OpenPyXL**: Excel file handling

## Setup and Installation

1. **Clone the Repository**:
   ```
   git clone [repository-url]
   cd [repository-name]
   ```

2. **Set Up a Virtual Environment** (recommended):
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install Dependencies**:
   ```
   pip install -r requirements.txt
   ```

4. **Configure Font Paths**:
   - Update the font file paths in `modules_external.py` to match your system's font locations.

5. **Run Experimental Data** (if applicable):
   - Use the dummy data generator provided to get experimental data to play with the demo.

## Running the Application

To start the Watts application, run:

```
python main.py
```

## Project Structure

- `main.py`: Entry point of the application
- `modules_external.py`: External module imports and configurations
- `settlement_api_dispatch_data.py`: Handles dispatch data processing
- `settlement_cs_performance_calculations.py`: Performs Connected Solutions performance calculations
- `settlement_cs_report_generation.py`: Generates reports for Connected Solutions
- `stylesheet.py`: Contains the CSS-like stylesheet for the application

## Key Components

### Watts Class

The main application class that sets up the GUI and manages the core functionality.

### CSPerformanceCalculator

Calculates performance metrics for Connected Solutions assets.

### CSReportGenerator

Generates detailed reports for assets and customers.

### Data Viewers

- `create_heatmap_settle_cs`: Creates a heatmap for performance analysis
- `create_geo_analysis_settle_cs`: Generates a geographical analysis of assets

## Usage Guide

1. **Data Import**: Use the "Pull Data" functionality to import the latest energy data.
2. **Event Selection**: Choose between Daily and Targeted events using the toggle buttons.
3. **Performance Calculation**: Select assets and use the "Calculate" button to process performance metrics.
4. **Report Generation**: Generate individual asset reports or customer reports using the respective buttons.
5. **Data Analysis**: Use the heatmap and geo-analysis tools for visual data exploration.

## Troubleshooting

- If you encounter issues with data loading, check the file paths and ensure all required files are present.
- For database connection issues, verify your Oracle client configuration and connection string.
- If visualizations are not rendering correctly, ensure all required libraries are correctly installed.

## Contributing

Contributions to the Watts project are welcome. Please follow these steps:

1. Fork the repository
2. Create a new branch (`git checkout -b feature-branch`)
3. Make your changes and commit (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin feature-branch`)
5. Create a new Pull Request

## MIT License

Copyright (c) 2024 Tyler Joe Sims

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Contact

For support or queries, please contact Tyler Joe Sims at tyler.jsims97@gmail.com.

## Sensitive Data Access:
The data in this repo is artificially generated data for demonstration uses only.
Feel free to replicate and reuse the generators however you wish.
