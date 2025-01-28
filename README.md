# GoogleMaps_DataExtractor

GoogleMaps_DataExtractor is a Python-based tool designed to automate the extraction of data from Google Maps. This project comprises two main components:

## Components

1. **Location Search & Link Generation**
    - Reads locations from a text file.
    - Searches for these locations on Google Maps.
    - Saves the resulting links to an Excel file named `search_results.xlsx`.

2. **Data Extraction**
    - Takes links from `search_results.xlsx`.
    - Extracts detailed information such as:
        - Title
        - Rating
        - Website
        - Phone number
        - User reviews with names and ratings
    - Stores the extracted data in another Excel file named `Output.xlsx`.
    - Tracks progress by marking each processed link with "Done" to ensure the process can resume if interrupted.

## Features

- Automates the process of searching locations and gathering data from Google Maps.
- Organizes the extracted data into an Excel file for easy analysis and utilization.
- Provides a mechanism to resume data extraction from where it left off in case of interruptions.

## How to Use

1. **Location Search & Link Generation:**
    - Place your locations in a text file (one per line).
    - Run the `maps_search_links.py` script to generate `search_results.xlsx` with Google Maps links.

2. **Data Extraction:**
    - Use the `Maps_data.py` script to extract detailed information from the links in `search_results.xlsx`.
    - The data will be saved in `Output.xlsx`.

## Requirements

- Python 3.x
- Selenium
- openpyxl
- WebDriver Manager

## Installation

Install the required packages using pip:
```bash
pip install selenium openpyxl webdriver_manager
