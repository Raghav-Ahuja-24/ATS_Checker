# ATS Checker

A Streamlit web application for searching and scoring candidates from Excel-based recruitment databases.

## Features

- **Multi-file Upload**: Upload multiple `.xlsx` candidate files or use the default local file
- **Multi-keyword Search**: Add multiple keywords with chips, search across all fields
- **ATS Score**: Automated Applicant Tracking System scoring based on keyword matches across resume data
- **Dynamic Column Visibility**: Show/hide any column from the Excel data
- **Experience & CTC Filters**: Range sliders for filtering by years of experience and CTC
- **Naukri Integration**: Login to Naukri recruiter portal to fetch actual resume data for enhanced scoring
- **Excel Export**: Download filtered results as styled Excel files
- **Dark Mode UI**: Premium dark-themed interface

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deployment

Deployed on Streamlit Community Cloud.

## Tech Stack

- Python 3.10+
- Streamlit
- openpyxl
- pandas
- BeautifulSoup4
- Selenium (for Naukri login)
