import streamlit as st
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import os
from resume import parse_resume, parse_resume_from_bytes, save_resume, load_resume, delete_resume, get_expiry_info

# Page Configuration
st.set_page_config(
    page_title="Naukri Job Scraper Pro",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern UI
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: 700;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        padding: 1rem 0;
    }
    .sub-header {
        text-align: center;
        color: #6c757d;
        font-size: 1.2rem;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        border-radius: 10px;
        padding: 0.75rem;
        border: none;
        font-size: 1.1rem;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #764ba2 0%, #667eea 100%);
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
    }
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        text-align: center;
    }
    .filter-section {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 1rem;
    }
    div[data-testid="stExpander"] {
        background: white;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .job-card {
        background: white;
        border-radius: 12px;
        padding: 1.2rem 1.2rem 0.8rem 1.2rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        border-left: 4px solid #667eea;
        margin-bottom: 1rem;
        height: 100%;
    }
    .job-card h4 { margin: 0 0 0.2rem 0; font-size: 1rem; color: #1a1a2e; }
    .job-card .company { font-weight: 600; color: #667eea; font-size: 0.9rem; }
    .job-card .meta { color: #6c757d; font-size: 0.8rem; margin: 0.4rem 0; }
    .job-card .desc { color: #444; font-size: 0.82rem; margin: 0.5rem 0 0.8rem 0; }
    .apply-btn-red button, .apply-btn-red button:hover, .apply-btn-red button:focus {
        background: #e74c3c !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
    }
    .apply-btn-green button, .apply-btn-green button:hover, .apply-btn-green button:focus {
        background: #27ae60 !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
    }
    /* hide the invisible state-tracking buttons */
    .track-btn { display: none !important; }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'scraped_data' not in st.session_state:
    st.session_state.scraped_data = pd.DataFrame()
if 'scraping_active' not in st.session_state:
    st.session_state.scraping_active = False
if 'scraping_stats' not in st.session_state:
    st.session_state.scraping_stats = {'total': 0, 'current_page': 0}
if 'resume_data' not in st.session_state:
    st.session_state.resume_data = None
if 'visited_jobs' not in st.session_state:
    st.session_state.visited_jobs = set()

OUTPUT_DIR = Path("output_jobs")
MAX_OUTPUT_FILES = 3

def get_output_files():
    """Return output xlsx files sorted oldest to newest by filename (timestamp embedded)"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    return sorted(OUTPUT_DIR.glob("naukri_jobs_*.xlsx"))

def load_existing_data():
    """Load the most recent output file"""
    files = get_output_files()
    if files:
        try:
            return pd.read_excel(files[-1])
        except:
            return pd.DataFrame()
    return pd.DataFrame()

def clear_output_files():
    """Delete all files in output_jobs/"""
    for f in get_output_files():
        f.unlink()

def save_to_excel(df, scraping_date):
    """Save scrape to output_jobs/ with timestamp filename, max 3 files"""
    if df.empty:
        return None
    OUTPUT_DIR.mkdir(exist_ok=True)
    df['Scraping_Date'] = scraping_date

    # Enforce max 3 files — delete oldest if needed
    files = get_output_files()
    while len(files) >= MAX_OUTPUT_FILES:
        files[0].unlink()
        files = get_output_files()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = OUTPUT_DIR / f"naukri_jobs_{timestamp}.xlsx"
    df.to_excel(file_path, index=False, engine='openpyxl')
    return file_path

def get_chrome_driver():
    """Initialize Chrome driver for both local (Windows) and cloud (Linux) environments"""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

    # On Linux (Streamlit Cloud) use system-installed chromium
    if os.name != 'nt':
        chrome_options.binary_location = "/usr/bin/chromium"
        service = Service("/usr/bin/chromedriver")
    else:
        service = Service(ChromeDriverManager().install())

    return webdriver.Chrome(service=service, options=chrome_options)

def scrape_naukri_jobs(job_titles, location, experience_min, experience_max, pages, progress_placeholder, table_placeholder, company_types=[]):
    """Scrape jobs with real-time updates"""
    print(f"\n{'='*60}")
    print(f"🚀 Starting Naukri Job Scraper")
    print(f"{'='*60}")
    print(f"📋 Job Titles: {', '.join(job_titles)}")
    print(f"📍 Location: {location}")
    print(f"💼 Experience: {experience_min}-{experience_max} years")
    print(f"🏢 Company Types: {', '.join(company_types) if company_types else 'All'}")
    print(f"📄 Pages to scrape: {pages}")
    print(f"{'='*60}\n")
    
    try:
        print("🔧 Initializing Chrome WebDriver...")
        driver = get_chrome_driver()
        print("✅ Chrome WebDriver initialized successfully\n")
    except Exception as e:
        print(f"❌ Failed to initialize Chrome driver: {e}")
        st.error(f"❌ Failed to initialize Chrome driver: {e}")
        return pd.DataFrame()
    
    all_jobs = []
    scraping_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Loop through each job title
    for job_title in job_titles:
        print(f"\n{'='*60}")
        print(f"🔍 Searching for: {job_title}")
        print(f"{'='*60}")
        
        for page in range(pages):
            if not st.session_state.scraping_active:
                break
                
            start = page * 20
            url = f"https://www.naukri.com/{job_title.replace(' ', '-')}-jobs-in-{location.lower()}-{start}"
            
            st.session_state.scraping_stats['current_page'] = page + 1
            print(f"\n🔍 Page {page + 1}/{pages}: {url}")
            progress_placeholder.info(f"🔄 Scraping {job_title} - Page {page + 1}/{pages}...")
            
            try:
                driver.get(url)
                print(f"⏳ Waiting for page to load...")
                time.sleep(3)
                
                # Take screenshot for debugging (only on first page)
                if page == 0:
                    try:
                        driver.save_screenshot(f"debug_page_{job_title.replace(' ', '_')}.png")
                        print(f"📸 Screenshot saved for debugging")
                    except:
                        pass
                
                wait = WebDriverWait(driver, 15)
                print(f"🔍 Looking for job cards with CSS selector...")
                wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "article.jobTuple, div.srp-jobtuple-wrapper")))
                
                jobs = driver.find_elements(By.CSS_SELECTOR, "article.jobTuple, div.srp-jobtuple-wrapper")
                print(f"📊 Found {len(jobs)} job elements on page")
                
                if len(jobs) == 0:
                    print(f"⚠️ No jobs found with primary selector, trying alternative...")
                    jobs = driver.find_elements(By.CSS_SELECTOR, "article, div[class*='job'], div[class*='tuple']")
                    print(f"🔄 Alternative selector found {len(jobs)} elements")
                
                jobs_extracted = 0
                for idx, job in enumerate(jobs):
                    try:
                        title = job.find_element(By.CSS_SELECTOR, "a.title").text.strip()
                        link = job.find_element(By.CSS_SELECTOR, "a.title").get_attribute("href")
                        company = job.find_element(By.CSS_SELECTOR, "a.comp-name").text.strip()
                        
                        location_elem = job.find_elements(By.CSS_SELECTOR, "span.locWdth")
                        location_text = location_elem[0].text.strip() if location_elem else "Not mentioned"
                        
                        exp_elem = job.find_elements(By.CSS_SELECTOR, "span.expwdth")
                        experience = exp_elem[0].text.strip() if exp_elem else "Not mentioned"
                        
                        desc_elem = job.find_elements(By.CSS_SELECTOR, "span.job-desc")
                        desc = desc_elem[0].text.strip() if desc_elem else "No description"
                        
                        posted_elem = job.find_elements(By.CSS_SELECTOR, "span.job-post-day")
                        posted = posted_elem[0].text.strip() if posted_elem else "N/A"
                        
                        tags_elements = job.find_elements(By.CSS_SELECTOR, "ul.tags-gt li")
                        tags = ', '.join([tag.text for tag in tags_elements]) if tags_elements else "N/A"
                        
                        # Detect company type from description or company name
                        company_type = "Not Specified"
                        desc_lower = desc.lower() + company.lower()
                        if any(word in desc_lower for word in ['product based', 'product company', 'saas']):
                            company_type = "Product Based"
                        elif any(word in desc_lower for word in ['service based', 'consulting', 'it services']):
                            company_type = "Service Based"
                        elif any(word in desc_lower for word in ['startup', 'early stage', 'funded']):
                            company_type = "Startup"
                        
                        job_data = {
                            "Title": title,
                            "Company": company,
                            "Company Type": company_type,
                            "Location": location_text,
                            "Experience": experience,
                            "Description": desc,
                            "Skills": tags,
                            "Posted": posted,
                            "Job URL": link
                        }
                        
                        # Filter by company type if specified
                        if not company_types or company_type in company_types or company_type == "Not Specified":
                            all_jobs.append(job_data)
                            jobs_extracted += 1
                            st.session_state.scraping_stats['total'] = len(all_jobs)

                            # Update table every 5 jobs to avoid flooding the websocket
                            if len(all_jobs) % 5 == 0:
                                temp_df = pd.DataFrame(all_jobs)
                                table_placeholder.dataframe(temp_df, width='stretch', height=400)
                        
                    except Exception as e:
                        print(f"⚠️ Failed to extract job #{idx}: {str(e)}")
                        continue
                
                print(f"✅ Page {page + 1} completed - Extracted {jobs_extracted}/{len(jobs)} jobs | Total so far: {len(all_jobs)}")
                progress_placeholder.success(f"✅ Page {page + 1} completed - {jobs_extracted} jobs extracted | Total: {len(all_jobs)}")
                # Always update table at end of each page
                if all_jobs:
                    table_placeholder.dataframe(pd.DataFrame(all_jobs), width='stretch', height=400)
                time.sleep(2)
                
            except Exception as e:
                print(f"❌ Error on page {page + 1}: {str(e)}")
                # Dump page source snippet to diagnose what Naukri returned
                try:
                    src = driver.page_source[:2000]
                    print(f"📄 Page source snippet:\n{src}")
                except:
                    pass
                progress_placeholder.error(f"❌ Error on page {page + 1}: {str(e)}")
                continue
    
    print("\n🔒 Closing browser...")
    driver.quit()
    
    if all_jobs:
        df = pd.DataFrame(all_jobs)
        print(f"\n{'='*60}")
        print(f"✅ Scraping Complete!")
        print(f"{'='*60}")
        print(f"📊 Total jobs scraped: {len(all_jobs)}")
        saved_path = save_to_excel(df, scraping_date)
        print(f"💾 Saved to: {saved_path}")
        print(f"✅ Data saved successfully!")
        print(f"{'='*60}\n")
        return df
    
    print("\n⚠️ No jobs found. Try different filters.\n")
    return pd.DataFrame()

# Header
st.markdown('<h1 class="main-header">🔍 Naukri Job Scraper Pro</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Real-time job scraping with advanced filters and analytics</p>', unsafe_allow_html=True)

# Sidebar Filters
with st.sidebar:
    st.markdown("### 🎯 Search Filters")
    
    # Resume Upload Section
    st.markdown("### 📄 Resume Auto-Fill")

    # On every load: check disk for a persisted resume and auto-load into session
    if st.session_state.resume_data is None:
        existing_path, _ = load_resume()
        if existing_path:
            resume_data = parse_resume(existing_path)
            if resume_data:
                st.session_state.resume_data = resume_data

    # Show expiry info if resume is on disk
    expiry_info = get_expiry_info()
    if expiry_info and not expiry_info.get("expired"):
        st.success(f"✅ Resume active: **{expiry_info['filename']}**")
        st.caption(f"⏳ Expires in: **{expiry_info['remaining_hours']}h {expiry_info['remaining_minutes']}m** — auto-deletes at {expiry_info['expiry_time']}")
        col_parse, col_del = st.columns(2)
        with col_parse:
            if st.button("🔍 Re-parse Resume"):
                existing_path, _ = load_resume()
                if existing_path:
                    with st.spinner("Parsing..."):
                        resume_data = parse_resume(existing_path)
                        if resume_data:
                            st.session_state.resume_data = resume_data
                            st.success("✅ Done!")
                            st.rerun()
        with col_del:
            if st.button("🗑️ Remove Resume"):
                delete_resume()
                clear_output_files()
                st.session_state.resume_data = None
                st.session_state.scraped_data = pd.DataFrame()
                st.session_state.visited_jobs = set()
                st.rerun()
    else:
        # Upload widget — only shown when no active resume
        uploaded_file = st.file_uploader(
            "Upload Resume (PDF / DOCX / TXT)",
            type=["pdf", "docx", "doc", "txt"],
            help="Only 1 resume at a time. Auto-deleted after 48 hours."
        )
        if uploaded_file:
            file_bytes = uploaded_file.read()
            with st.spinner("Saving & parsing resume..."):
                saved_path = save_resume(file_bytes, uploaded_file.name)
                resume_data = parse_resume_from_bytes(file_bytes, uploaded_file.name)
                if resume_data:
                    st.session_state.resume_data = resume_data
                    st.success("✅ Resume uploaded & parsed!")
                    st.rerun()
                else:
                    st.error("❌ Could not parse resume. Check file format.")

    # Show parsed resume info
    if st.session_state.resume_data:
        with st.expander("📋 Parsed Resume Info", expanded=False):
            rd = st.session_state.resume_data
            st.write(f"**Experience:** {rd['experience']} years")
            st.write(f"**Skills Match:** {len(rd['job_titles'])} job titles")
            st.write(f"**Location:** {rd['location']}")
            if rd['company_types']:
                st.write(f"**Preferences:** {', '.join(rd['company_types'])}")
    
    st.markdown("---")
    st.markdown("### 🔧 Job Filters")
    
    with st.expander("📋 Job Details", expanded=True):
        # Job titles dropdown with multiple selection
        job_titles_list = [
            "Web Developer", "Data Analyst", "Python Developer", "Full Stack Developer",
            "Frontend Developer", "Backend Developer", "Software Engineer", "DevOps Engineer",
            "Data Scientist", "Machine Learning Engineer", "Java Developer", "React Developer",
            "Node.js Developer", "Angular Developer", "UI/UX Designer", "Product Manager",
            "Business Analyst", "QA Engineer", "Cloud Engineer", "Cybersecurity Analyst"
        ]
        
        # Auto-fill from resume or use default
        default_jobs = st.session_state.resume_data['job_titles'] if st.session_state.resume_data else ["Web Developer"]
        
        selected_jobs = st.multiselect(
            "Job Titles (Select Multiple)",
            options=job_titles_list,
            default=default_jobs,
            help="Select one or more job titles"
        )
        
        # Auto-fill company types from resume
        default_company_types = st.session_state.resume_data['company_types'] if st.session_state.resume_data else []
        
        # Company type filter
        company_types = st.multiselect(
            "Company Type",
            options=["Product Based", "Service Based", "Startup"],
            default=default_company_types,
            help="Filter by company type (optional)"
        )
        
        # Auto-fill location from resume
        default_location = st.session_state.resume_data['location'] if st.session_state.resume_data else "India"
        
        location = st.text_input("Location", value=default_location, help="e.g., Bangalore, Mumbai, India")
    
    with st.expander("💼 Experience Range", expanded=True):
        # Auto-fill experience from resume
        if st.session_state.resume_data:
            resume_exp = st.session_state.resume_data['experience']
            default_min = max(0, resume_exp - 2)
            default_max = resume_exp + 3
        else:
            default_min = 0
            default_max = 10
        
        col1, col2 = st.columns(2)
        with col1:
            exp_min = st.number_input("Min (years)", min_value=0, max_value=30, value=default_min)
        with col2:
            exp_max = st.number_input("Max (years)", min_value=0, max_value=30, value=default_max)
    
    with st.expander("📄 Scraping Settings", expanded=True):
        num_pages = st.slider("Number of Pages", min_value=1, max_value=50, value=5, help="Each page contains ~20 jobs")
        st.info(f"📊 Estimated jobs: ~{num_pages * 20}")
    
    st.markdown("---")
    
    # Scrape Button
    if not st.session_state.scraping_active:
        if st.button("🚀 Start Scraping", type="primary"):
            if not selected_jobs:
                st.error("⚠️ Please select at least one job title")
            elif st.session_state.resume_data is None:
                st.toast("📄 Please upload your resume first to auto-fill filters!", icon="⚠️")
            else:
                st.session_state.scraping_active = True
                st.session_state.scraping_stats = {'total': 0, 'current_page': 0}
                st.rerun()
    else:
        if st.button("⏹️ Stop Scraping", type="secondary"):
            st.session_state.scraping_active = False
            st.rerun()
    
    # Load Historical Data
    st.markdown("---")
    if st.button("📂 Load Historical Data"):
        historical_data = load_existing_data()
        if not historical_data.empty:
            st.session_state.scraped_data = historical_data
            st.success(f"✅ Loaded {len(historical_data)} jobs")
        else:
            st.warning("No historical data found")

# Main Content Area
tab1, tab2, tab3, tab4 = st.tabs(["🔍 Live Scraping", "📊 Analytics", "📥 Data Export", "💼 Job Cards"])

with tab1:
    if st.session_state.scraping_active:
        st.markdown("### 🔄 Scraping in Progress...")
        
        # Progress section
        progress_placeholder = st.empty()
        
        # Stats
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Jobs Scraped", st.session_state.scraping_stats['total'])
        with col2:
            st.metric("📖 Current Page", st.session_state.scraping_stats['current_page'])
        with col3:
            st.metric("🎯 Target Pages", num_pages)
        
        st.markdown("---")
        st.markdown("### 📋 Live Job Listings")
        table_placeholder = st.empty()
        
        # Start scraping
        scraped_df = scrape_naukri_jobs(
            selected_jobs, location, exp_min, exp_max, num_pages,
            progress_placeholder, table_placeholder, company_types
        )
        
        st.session_state.scraped_data = scraped_df
        st.session_state.scraping_active = False
        
        if not scraped_df.empty:
            st.success(f"✅ Scraping completed! Total jobs: {len(scraped_df)}")
            st.balloons()
        else:
            st.error("❌ No jobs found. Try different filters.")
    
    else:
        if st.session_state.scraped_data.empty:
            st.info("👈 Configure filters in the sidebar and click 'Start Scraping' to begin")
            
            # Show sample preview
            st.markdown("### 📝 Preview")
            st.markdown("""
            **What you'll get:**
            - Job Title & Company
            - Location & Experience Required
            - Skills & Job Description
            - Posted Date & Direct Job URL
            - Real-time updates during scraping
            - Auto-save to Excel with timestamps
            """)
        else:
            st.markdown("### 📋 Scraped Job Listings")
            st.dataframe(st.session_state.scraped_data, width='stretch', height=500)

with tab2:
    if not st.session_state.scraped_data.empty:
        df = st.session_state.scraped_data
        
        # Summary Stats
        st.markdown("### 📊 Summary Statistics")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📄 Total Jobs", len(df))
        with col2:
            st.metric("🏢 Companies", df['Company'].nunique())
        with col3:
            st.metric("📍 Locations", df['Location'].nunique())
        with col4:
            if 'Scraping_Date' in df.columns:
                st.metric("🕒 Last Updated", df['Scraping_Date'].iloc[-1][:10])
            else:
                st.metric("🕒 Session", "Current")
        
        st.markdown("---")
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 🏢 Top 10 Hiring Companies")
            top_companies = df['Company'].value_counts().head(10).reset_index()
            top_companies.columns = ['Company', 'Count']
            fig1 = px.bar(top_companies, x='Count', y='Company', orientation='h',
                         color='Count', color_continuous_scale='Viridis')
            fig1.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig1, width='stretch')
        
        with col2:
            st.markdown("#### 📍 Top 10 Job Locations")
            top_locations = df['Location'].value_counts().head(10).reset_index()
            top_locations.columns = ['Location', 'Count']
            fig2 = px.bar(top_locations, x='Count', y='Location', orientation='h',
                         color='Count', color_continuous_scale='Plasma')
            fig2.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig2, width='stretch')
        
        # Posted Time Distribution
        st.markdown("#### ⏳ Job Posting Timeline")
        posted_counts = df['Posted'].value_counts().reset_index()
        posted_counts.columns = ['Posted', 'Count']
        fig3 = px.pie(posted_counts, names='Posted', values='Count', hole=0.4,
                     color_discrete_sequence=px.colors.qualitative.Set3)
        fig3.update_layout(height=400)
        st.plotly_chart(fig3, width='stretch')
        
    else:
        st.info("📊 Analytics will appear here after scraping jobs")

with tab3:
    output_files = get_output_files()
    if not output_files:
        st.info("📥 No output files yet. Start scraping to generate data.")
    else:
        st.markdown("### 📥 Saved Scrape Files")
        st.caption(f"💾 {len(output_files)}/{MAX_OUTPUT_FILES} files stored — oldest auto-removed on next scrape")
        st.markdown("---")

        for i, fpath in enumerate(reversed(output_files)):  # newest first
            try:
                fdf = pd.read_excel(fpath)
            except:
                fdf = pd.DataFrame()

            label = ("🟢 Latest" if i == 0 else f"🕓 Older #{i}")
            scrape_date = fdf['Scraping_Date'].iloc[0] if 'Scraping_Date' in fdf.columns and not fdf.empty else "Unknown"

            with st.expander(f"{label} — {fpath.name}  |  {len(fdf)} jobs  |  Scraped: {scrape_date}", expanded=(i == 0)):
                ec1, ec2 = st.columns(2)
                with ec1:
                    csv = fdf.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📄 Download as CSV",
                        data=csv,
                        file_name=fpath.stem + ".csv",
                        mime="text/csv",
                        width='stretch',
                        key=f"csv_{fpath.name}"
                    )
                with ec2:
                    with open(fpath, 'rb') as f:
                        st.download_button(
                            label="📊 Download Excel",
                            data=f,
                            file_name=fpath.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch',
                            key=f"xlsx_{fpath.name}"
                        )

                if not fdf.empty:
                    mc1, mc2, mc3 = st.columns(3)
                    mc1.metric("📄 Total Jobs", len(fdf))
                    mc2.metric("🏢 Companies", fdf['Company'].nunique() if 'Company' in fdf.columns else 'N/A')
                    mc3.metric("📍 Locations", fdf['Location'].nunique() if 'Location' in fdf.columns else 'N/A')
                    st.dataframe(fdf, width='stretch', height=350)
                else:
                    st.warning("Could not read file data.")

with tab4:
    if st.session_state.scraped_data.empty:
        st.info("💼 Job cards will appear here after scraping or loading historical data")
    else:
        df_cards = st.session_state.scraped_data.copy()

        # ── Filters ──────────────────────────────────────────────
        st.markdown("### 🔎 Filter Jobs")
        fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 2])
        with fc1:
            all_titles = sorted(df_cards['Title'].dropna().unique().tolist())
            sel_title = st.multiselect("Job Title", all_titles, key="card_title")
        with fc2:
            all_companies = sorted(df_cards['Company'].dropna().unique().tolist())
            sel_company = st.multiselect("Company", all_companies, key="card_company")
        with fc3:
            all_locs = sorted(df_cards['Location'].dropna().unique().tolist())
            sel_loc = st.multiselect("Location", all_locs, key="card_loc")
        with fc4:
            show_visited = st.radio("Status", ["All", "Not Applied", "Applied"], horizontal=True, key="card_status")

        # Experience range slider — parse both bounds from strings like "2-5 Yrs", "Fresher", "10+ Yrs"
        import re as _re
        def _parse_exp(val):
            """Returns (min, max) tuple. Fresher -> (0,0). '3-8 Yrs' -> (3,8). '10+ Yrs' -> (10,30)."""
            s = str(val).strip().lower()
            if not val or s in ('not mentioned', 'n/a', 'nan', ''):
                return None
            if 'fresher' in s:
                return (0, 0)
            nums = list(map(int, _re.findall(r'\d+', s)))
            if len(nums) == 0:
                return None
            if len(nums) == 1:
                return (nums[0], nums[0])  # e.g. "10+ Yrs" -> (10, 10)
            return (nums[0], nums[1])      # e.g. "3-8 Yrs" -> (3, 8)

        df_cards['_exp_parsed'] = df_cards['Experience'].apply(_parse_exp)
        valid_exp = df_cards['_exp_parsed'].dropna()

        if not valid_exp.empty:
            all_mins = [v[0] for v in valid_exp]
            all_maxs = [v[1] for v in valid_exp]
            exp_floor = int(min(all_mins))
            exp_ceil  = int(max(all_maxs))

            if exp_floor < exp_ceil:
                fe1, fe2 = st.columns([3, 1])
                with fe1:
                    exp_range = st.slider(
                        "💼 Experience Range (years)",
                        min_value=exp_floor, max_value=exp_ceil,
                        value=(exp_floor, exp_ceil), key="card_exp"
                    )
                with fe2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    include_unspecified = st.checkbox("Include unspecified", value=True, key="card_exp_unspec")

                def _exp_passes(parsed):
                    if parsed is None:
                        return include_unspecified
                    # keep job if its range overlaps with selected range
                    return parsed[0] <= exp_range[1] and parsed[1] >= exp_range[0]

                df_cards = df_cards[df_cards['_exp_parsed'].apply(_exp_passes)]

        df_cards = df_cards.drop(columns=['_exp_parsed'])

        if sel_title:
            df_cards = df_cards[df_cards['Title'].isin(sel_title)]
        if sel_company:
            df_cards = df_cards[df_cards['Company'].isin(sel_company)]
        if sel_loc:
            df_cards = df_cards[df_cards['Location'].isin(sel_loc)]
        if show_visited == "Applied":
            df_cards = df_cards[df_cards['Job URL'].isin(st.session_state.visited_jobs)]
        elif show_visited == "Not Applied":
            df_cards = df_cards[~df_cards['Job URL'].isin(st.session_state.visited_jobs)]

        st.caption(f"Showing **{len(df_cards)}** jobs")
        st.markdown("---")

        # ── Cards grid (3 columns) ────────────────────────────────
        COLS = 3
        rows = [df_cards.iloc[i:i+COLS] for i in range(0, len(df_cards), COLS)]

        for row in rows:
            cols = st.columns(COLS)
            for col, (row_idx, job) in zip(cols, row.iterrows()):
                url = job.get('Job URL', '#') or '#'
                visited = url in st.session_state.visited_jobs
                desc = str(job.get('Description', ''))
                short_desc = (desc[:50] + '...') if len(desc) > 50 else desc
                btn_color = "#27ae60" if visited else "#e74c3c"
                btn_label = "✅ Applied" if visited else "🔗 Apply Now"

                with col:
                    st.markdown(f"""
                    <div class="job-card">
                        <h4>{job.get('Title', 'N/A')}</h4>
                        <div class="company">🏢 {job.get('Company', 'N/A')}</div>
                        <div class="meta">📍 {job.get('Location', 'N/A')} &nbsp;|&nbsp; 💼 {job.get('Experience', 'N/A')}</div>
                        <div class="desc">{short_desc}</div>
                        <a href="{url}" target="_blank"
                           style="display:block;text-align:center;padding:0.45rem 0;border-radius:8px;
                                  font-weight:600;font-size:0.88rem;text-decoration:none;
                                  color:white;background:{btn_color};margin-top:0.5rem;"
                           onclick="document.getElementById('track_{row_idx}').click();"
                        >{btn_label}</a>
                    </div>
                    """, unsafe_allow_html=True)
                    st.markdown('<div class="track-btn">', unsafe_allow_html=True)
                    if st.button("t", key=f"track_{row_idx}"):
                        st.session_state.visited_jobs.add(url)
                        st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #6c757d; padding: 1rem;'>
    <p>🔍 Naukri Job Scraper Pro | Built with Streamlit & Selenium</p>
    <p style='font-size: 0.9rem;'>💡 Tip: Use filters to narrow down your search and get better results</p>
</div>
""", unsafe_allow_html=True)
