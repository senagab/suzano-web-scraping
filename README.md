# Automation with Web Scraping for Suzano

<p align="left">
    <img src="https://img.shields.io/badge/selenium-v4.28.1-green?logo=selenium&labelColor=white" alt="Selenium version">
    <img src="https://img.shields.io/badge/python-v3.13.2-blue?logo=python&labelColor=white" alt=" version">
</p>

This project automates the process of filling out a web form using Selenium and data from an Excel spreadsheet. A simple Tkinter-based GUI allows users to trigger the automation with a button click.

## Features
- Reads data from an Excel file (`banco.xlsx`).
- Launches a Chrome browser in fullscreen mode.
- Automatically fills out form fields on a specified webpage.
- Submits the form for each row of data.
- Displays success or error messages upon completion.
- Provides a GUI button for easy execution.

## Prerequisites
Ensure you have the following installed:
- Python 3.x
- Google Chrome browser
- Chrome WebDriver (compatible with your Chrome version)
- Required Python packages:
  ```sh
  pip install selenium openpyxl tk
  ```

## Setup
1. Download the Chrome WebDriver from [ChromeDriver](https://sites.google.com/chromium.org/driver/) and place it in your project directory.
2. Update the `chromedriver.exe` path in the script if necessary.
3. Ensure `banco.xlsx` exists in the specified path.

## Usage
1. Run the script:
   ```sh
   python index.py
   ```
2. Click the "Start Automation" button in the GUI.
3. The browser will open in fullscreen and begin filling out forms.
4. Wait for the process to complete and check the console for messages.

## Notes
- Ensure the webpage elements (`ID`s) match the form fields being automated.
- The script stops processing when it encounters an empty row in the Excel sheet.
