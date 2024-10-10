# Mass Email Lookup
This repository contains an Email Lookup Tool that leverages the Perplexity API to retrieve contact information such as email addresses, phone numbers, and sources for individuals from an Excel file. The GUI for the tool is built using PySide6 for a modern, frameless interface with dark mode support, and it includes API key storage encryption.

## Installation
* Clone the repository using `git clone`
* To install the required dependencies, run the following command in your terminal:

        pip install -r requirements.txt

## Usage
### Running the Application
* You can start the application by running:

        python ui_test.py

* Choose an Excel file `(.xlsx)` that contains the following columns:
  * `FirstName`
  * `LastName`
  * `Title`
  * `Organization`
* Enter Perplexity API Key:
  * You need to input a valid Perplexity API Key in the field provided. The key can be saved securely for future use. Click on `Load API Key` after entering the key
* Start Email Lookup:
  * Once an Excel file is selected and the API key is entered and loaded, click "Start Email Lookup" to process the data. The tool will retrieve contact details and save the output to a new Excel file in the same directory from which the original was accessed.

### Features
* *Excel File Processing*: Reads from an Excel file and outputs the results to a new Excel file.
