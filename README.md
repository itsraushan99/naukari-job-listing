# 🔍 Naukri Job Scraper Pro

A real-time job scraping dashboard built with **Streamlit** and **Selenium** that scrapes jobs from [Naukri.com](https://www.naukri.com), auto-fills filters from your resume, and lets you track applications — all in one place.

---

## 📁 Project Structure

```
naukari-scrapper-dashboard_v2/
├── app.py                        # Main Streamlit application
├── requirements.txt              # Python dependencies
├── output_jobs/                  # Auto-generated Excel scrape files (max 3)
│   └── naukri_jobs_<timestamp>.xlsx
└── resume/
    ├── __init__.py
    ├── resume_manager.py         # Resume upload, storage & expiry logic
    ├── resume_parser.py          # PDF / DOCX / TXT text extraction & parsing
    └── uploaded_resumes/
        ├── <your_resume>.pdf
        └── resume_meta.json      # Stores filename & upload timestamp
```

---

## ✨ Features

### 📄 Resume Auto-Fill
- Upload your resume (PDF, DOCX, DOC, TXT)
- The app automatically parses it and pre-fills:
  - Job titles (matched from skills/keywords)
  - Location (detected from Indian cities)
  - Experience range
  - Company type preference (Product / Service / Startup)
- Resume is stored locally and **auto-deleted after 48 hours**
- Only one resume is stored at a time

### 🚀 Real-Time Job Scraping
- Scrapes Naukri.com using headless Chrome (via Selenium)
- Supports **multiple job titles** in a single run
- Filters: location, experience range, pages (1–50, ~20 jobs/page), company type
- Live table updates as jobs are found
- Stop scraping at any time

### 💾 Output File Management
- Results auto-saved to `output_jobs/` as timestamped `.xlsx` files
- Maximum **3 files** stored — oldest is removed automatically on the next scrape
- Load any previous scrape via the "Load Historical Data" button

### 📊 Analytics Dashboard
- Top 10 hiring companies (bar chart)
- Top 10 job locations (bar chart)
- Job posting timeline (donut chart)
- Summary stats: total jobs, unique companies, unique locations, last updated

### 💼 Job Cards View
- Card-based layout (3 columns) with title, company, location, experience, description
- Filter cards by title, company, location, experience range, and application status
- Click "Apply Now" to open the job link — card turns green and is marked as **Applied**
- Applied/Not Applied status tracked within the session

### 📥 Data Export
- Download any saved scrape as **CSV** or **Excel (.xlsx)**
- View job count, companies, and locations per file before downloading

---

## 🚀 Getting Started

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the App

```bash
streamlit run app.py
```

### 3. Use the App

1. Upload your resume in the sidebar — filters auto-fill
2. Adjust job titles, location, experience, and pages if needed
3. Click **🚀 Start Scraping**
4. Watch jobs populate in real time
5. Explore **Analytics**, **Job Cards**, and **Data Export** tabs

---

## 📦 Dependencies

| Package            | Purpose                              |
|--------------------|--------------------------------------|
| `streamlit`        | Web UI framework                     |
| `selenium`         | Browser automation for scraping      |
| `webdriver-manager`| Auto-downloads matching ChromeDriver |
| `pandas`           | Data handling                        |
| `plotly`           | Interactive charts                   |
| `openpyxl`         | Excel file read/write                |
| `PyPDF2`           | PDF text extraction                  |
| `python-docx`      | DOCX text extraction                 |

---

## 🔧 How Resume Parsing Works

The parser (`resume_parser.py`) extracts text from your resume and uses keyword matching to detect:

| Field           | Method                                                        |
|-----------------|---------------------------------------------------------------|
| Experience      | Regex patterns like `"3+ years of experience"`               |
| Job Titles      | Keyword-to-title mapping (e.g. `react` → Frontend Developer) |
| Location        | Matches against a list of major Indian cities                 |
| Company Type    | Keywords like `"product based"`, `"startup"`, `"consulting"`  |

---

## 📊 Scraped Data Fields

| Field          | Description                          |
|----------------|--------------------------------------|
| Title          | Job position name                    |
| Company        | Hiring company                       |
| Company Type   | Product Based / Service Based / Startup / Not Specified |
| Location       | Job location(s)                      |
| Experience     | Required experience                  |
| Description    | Short job description                |
| Skills         | Required skills/tags                 |
| Posted         | When the job was posted              |
| Job URL        | Direct link to the job listing       |
| Scraping_Date  | Timestamp when data was scraped      |

---

## 🛠️ Troubleshooting

**ChromeDriver not found**
- Make sure Google Chrome is installed
- `webdriver-manager` handles driver download automatically on first run

**No jobs found**
- Try broader filters (e.g. location = `India`, fewer pages)
- Naukri.com may throttle rapid requests — wait a few minutes between runs

**Excel file locked**
- Close the file in Excel before starting a new scrape
- Use CSV export as an alternative

**Resume not parsing correctly**
- Ensure the resume is text-based (not a scanned image PDF)
- Try converting to `.txt` or `.docx` for better results

---

## ⚠️ Notes

- Application tracking (Applied / Not Applied) is **session-only** — it resets on page refresh
- Scraping Naukri.com is subject to their terms of service — use responsibly
- Resume files are stored locally in `resume/uploaded_resumes/` and deleted after 48 hours

---

*Built with ❤️ using Streamlit, Selenium, and Plotly*
