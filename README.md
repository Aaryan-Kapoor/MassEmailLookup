# Mass Email Lookup with an easy-to-use UI
This repository contains an Email Lookup Tool that leverages the Perplexity API to retrieve contact information such as email addresses, phone numbers, and sources for individuals from an Excel file. The GUI for the tool is built using PySide6 for a modern, frameless interface with dark mode support, and it includes API key storage encryption.

## Installation
* Clone the repository using `git clone https://github.com/Aaryan-Kapoor/MassEmailLookup`
* To install the required dependencies, run the following command in your terminal:

        pip install -r requirements.txt

## Usage
### Running the Application
* You can start the application by running:

        python ui_test.py

* Choose an Excel file `.xlsx` that contains the following columns:
  * `FirstName`
  * `LastName`
  * `Title`
  * `Organization`
* Enter Perplexity API Key:
  * You need to input a valid Perplexity API Key in the field provided. The key can be saved securely for future use.
  * Click on `Save API Key` to save it if needed or
  * Click on `Load API Key` after entering the key. The API key will be encrypted and stored safely.
* Start Email Lookup:
  * Once an Excel file is selected and the API key is entered and loaded, click "Start Email Lookup" to process the data. The tool will retrieve contact details and save the output to a new Excel file in the same directory from which the original was accessed.

### Features
* **Excel File Processing**: Reads from an Excel file and outputs the results to a new Excel file.
* **Perplexity API Integration**: Fetches the most accurate contact details using OpenAI’s API.
* **Secure API Key Storage**: API keys are encrypted and securely stored locally.
* **Progress Bar and Notifications**: The progress of the lookup process is shown in a progress bar, and a message is displayed once processing is complete.
* **Custom GUI**: A modern, frameless window with hover effects and dark mode.

## Contributions are Welcome!
