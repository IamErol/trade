# TN VED Processor

An automated tool for extracting, processing, and translating duty rates from Uzbek legal documents and web sources.

## Features
- **Document Parsing**: Extracts TN VED codes and descriptions from `.docx` files.
- **Automated Rate Fetching**: Scrapes current duty rates from Lex.uz.
- **Excel Export**: Generates a formatted `.xlsx` file with the processed data.

## Getting Started

### Prerequisites
- Python 3.8+
- Django

### Installation

1. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```
   
2. Clone the repository and navigate to the project directory:
   ```bash
   cd trade_web
   ```

3. Run the development server:
   ```bash
   python manage.py runserver
   ```

### Usage
1. Open your browser and go to `http://127.0.0.1:8000/`.
2. Enter the **Duty Rates URL** (default is usually correct for current legislation).
3. Upload your **TN VED document** (`.docx` format).
4. Click **Start Processing**.
5. Wait for the processing to complete and the Excel file will download automatically.
