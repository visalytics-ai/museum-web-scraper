# museum-web-scraper
This project scrapes the Met Museums website for images and descriptions of swords to build a database for a machine learning algorithm to identify the place and decade of origin of swords based on construction and stylistic traits. This scraper captures details not in the Met API. Check this space for updates on the machine learning progress.

______________
# Museum Sword Corpus Scraper

A scalable Python-based data pipeline for extracting structured sword metadata and imagery from major museum collections.

This project was built to construct a research-grade dataset of swords from institutional museum catalogs (beginning with the Metropolitan Museum of Art). It integrates API access, browser automation, DOM parsing, data cleaning, image management, and large-batch processing into a resilient scraping framework.

---

## 📌 Project Objective

The goal of this project was to:

- Programmatically query museum search endpoints for all sword objects
- Extract structured API metadata
- Render JavaScript-driven object pages to capture fields not exposed via API
- Extract tabbed sections such as inscriptions, provenance, and references
- Download and organize primary and additional images into per-object folders
- Persist cleaned metadata and image references in a structured format suitable for analysis
- Support long-running scrapes (3,000–5,000 objects) with resume and checkpoint logic

---

## 🏗 Architecture Overview

### Core Technologies

- Python
- Playwright (headless Chromium rendering)
- Requests
- BeautifulSoup
- OpenPyXL (initial output format)
- Structured file system storage for images

---

## 🔄 High-Level Workflow

1. Query museum search API to retrieve object IDs.
2. For each object:
   - Fetch structured JSON via API (if available)
   - Render the webpage via Playwright to execute JavaScript
   - Programmatically click dynamic tabs
   - Extract text content from rendered DOM
   - Extract long-form descriptions using layered fallbacks
   - Extract primary and additional image URLs
   - Download all images into:

     ```
     downloaded_images/<objectID>/<objectID>_#.jpg
     ```

   - Clean text for Excel-safe characters
   - Write structured output with checkpoint saving
3. Support resume via configurable offset.

---

## ⚙️ Major Engineering Challenges & Solutions

### 1️⃣ API Incompleteness

The museum API did not expose key fields including:

- Inscriptions  
- Provenance  
- References  
- Long-form descriptions  

**Solution:**  
Integrated Playwright to render the live webpage and extract dynamically injected DOM content.

---

### 2️⃣ JavaScript-Rendered Tab Content

Important metadata appeared only after clicking tab elements.

**Solution:**  
Automated tab clicks via Playwright, waited for DOM updates, and extracted the updated content container.

---

### 3️⃣ Long Description Fragmentation

Long descriptions were not stored in a single predictable element. They appeared in:

- Custom React components  
- Expandable read-more wrappers  
- Meta description tags  
- Inconsistent span structures  

**Solution:**  
Implemented a hierarchical fallback extractor:

1. Preferred React span containers  
2. Longest valid descriptive span  
3. Meta description  
4. Open Graph description  
5. Paragraph fallback  

---

### 4️⃣ Illegal Excel Characters

Certain inscriptions contained control characters (ASCII 0x00–0x1F), causing:

```
openpyxl.utils.exceptions.IllegalCharacterError
```

**Solution:**  
Sanitized every string before writing:

```python
ILLEGAL_CHARS_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")
```

Applied uniformly across all row values.

---

### 5️⃣ Large-Scale Batch Stability

Full dataset contained ~4,000 objects.

Challenges:
- Network timeouts  
- Page load failures  
- System sleep interruption  
- Excel file bloat  
- Data loss risk  

**Solutions implemented:**

- Per-object exception handling  
- Navigation timeout protection  
- Checkpoint saving every N objects  
- Resume offset configuration  
- Only embedding primary thumbnails to reduce file size  
- Storing full-resolution images externally  

---

### 6️⃣ Image Management at Scale

Embedding all images in Excel produced multi-gigabyte files and corruption warnings.

**Solution:**

- Embed only the primary thumbnail  
- Store all images on disk in per-object folders  
- Store URLs and local paths in dataset  

---

## 📊 Results

The pipeline successfully scraped the full Metropolitan Museum Arms & Armor sword corpus (~4,000 objects) with:

- Complete structured metadata  
- Long descriptions  
- Tabbed content  
- All image URLs  
- Organized image archive  
- Clean, corruption-free output  

The architecture is modular and extensible to other museum systems (Smithsonian, V&A, British Museum, etc.).

---

## 🔁 Resume Support

The scraper includes a configurable:

```
START_OFFSET
```

This allows recovery from interruption without restarting from scratch.

---

## 📂 Project Structure

```
museum-sword-scraper/
│
├── met_swords_full_scrape.py
├── downloaded_images/
│   └── <objectID>/
│       ├── 12345_1.jpg
│       ├── 12345_2.jpg
│       └── ...
├── README.md
└── requirements.txt
```

---

## 🚀 Future Directions

- Replace Excel output with SQLite or Parquet for scalable analytics  
- Build multi-museum ingestion architecture  
- Add structured relational schema for objects/images  
- Integrate lightweight dataset browser  

---

## 💡 What This Project Demonstrates

- Real-world scraping beyond static HTML  
- Robust automation against dynamic React applications  
- Scalable data engineering practices  
- Defensive error handling in long-running jobs  
- Production-aware design (checkpointing, resume logic, file hygiene)  

This project is designed as a reusable museum corpus ingestion framework rather than a one-off script.
