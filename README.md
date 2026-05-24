# Web-Scrapper-Using-Python

# Playwright B2B Data Scraper Pipeline

A robust, asynchronous web scraping pipeline built with Python, Playwright, and BeautifulSoup4. This system mimics human browsing behavior to bypass basic bot-detection mechanisms, extracts structured data fields dynamically, and exports the data into clean CSV formats ready for downstream processing or machine learning pipelines.

---

## 🛠️ Tech Stack & Architecture

* **Runtime Environment:** Python 3.10+
* **Browser Automation:** Playwright (Chromium Engine)
* **HTML Parsing Engine:** BeautifulSoup4
* **Data Serialization:** Pandas
* **Security Evading:** Fake-UserAgent (Dynamic Header Rotation)

The scraper utilizes an asynchronous architecture to efficiently handle network latency. By launching a headless browser instance, it fully renders JavaScript-heavy applications before extracting the underlying page source.

---

## 🚀 Getting Started

### 1. Prerequisites
Ensure you have Python installed. If you are using an **Anaconda/Conda** environment, make sure your environment is activated (`(base)` or your custom env) before running the installations.

### 2. Installation
Install the required dependencies directly to your active environment:

```bash
pip install playwright beautifulsoup4 pandas fake-useragent
