# Cascade

### Turn Raw Data Into Insight

Cascade is a Django web application that takes a raw dataset, cleans it, builds charts and pivot tables from it, writes the insights, and exports the whole thing as a PDF report.

You stay in control of the analysis. Cascade lists every change it makes while cleaning, never fills in missing values for you, and lets you edit every written insight, summary and recommendation before you export.

## Features

* **Dataset Upload**
  * Supports `.csv`, `.xlsx`, `.xls`, `.json`, `.jsonl` and `.ndjson` files, up to 5 MB
  * JSON can be a list of records, JSON Lines, or an object wrapping a list. Nested objects become columns named `parent_child`
  * Excel files read the first sheet only
  * Your original file is never modified

* **Data Cleaning**
  * Fixes automatically and lists every change in a cleaning report:
    * Removes completely empty rows and exact duplicate rows
    * Trims text and collapses repeated spaces
    * Treats `NA`, `N/A`, `null`, `-` and blanks as missing
    * Converts numeric-looking text to numbers (currency symbols and commas removed, the symbol is remembered for display)
    * Lowercases column names and replaces spaces with underscores
    * Merges category spellings that differ only by capitalization
    * Standardizes dates to `YYYY-MM-DD` (`YYYY-MM` for month-level dates) and converts Excel or Unix date numbers in date-named columns
  * Flags possible outliers and near-duplicate rows for review, without changing them
  * **You decide what happens to rows with missing values:** remove them, or review them one by one and type in a value or remove the row. Typed values are validated against the column, and nothing is applied until you confirm
  * Download the cleaned data as a CSV

* **Key Metrics and Filters**
  * Headline number cards (up to 12) with totals, averages, medians, minimums, maximums, counts, distinct counts and most-common values
  * Filters (up to 12) by number range, date range, or text values, applied to every metric, chart, pivot table and the PDF
  * With a filter on, each metric also shows the figure for the full dataset

* **Pivot Tables**
  * One row field, an optional column field and one measure (Sum, Average, Median, Minimum, Maximum, Count or Distinct count)
  * Group dates by day, week, month, quarter or year
  * Rank column, % of total, and Top N / Bottom N
  * Download any pivot table as a CSV

* **Data Visualization**
  * Chart types: Auto, Bar, Horizontal Bar, Pie, Line, Area, Scatter, Histogram, Box Plot and Correlation Heatmap
  * Date grouping for line, area and bar charts, with Sum, Average, Count, Maximum or Minimum
  * Sort, and Top N / Bottom N, for bar, horizontal bar and pie charts
  * Calculated values (ƒx), such as profit margin as `Profit ÷ Revenue × 100`
  * Reorder charts and pivot tables by dragging or with the ▲ ▼ buttons

* **Insights and Recommendations**
  * Every chart and pivot table gets a written insight based on the data actually plotted
  * An overall summary and a list of recommendations are generated from your metrics and charts
  * All written text can be edited and saved. Cascade warns you before it replaces anything you wrote

* **Save and Resume Progress**
  * Save your work as a `.cascade` file (up to 15 MB) and import it later, on any device, without uploading or cleaning again
  * The file holds the cleaned data, charts, pivots, filters, metrics, page order and your edited notes
  * Imports are checked against a checksum and validated, so a damaged file never changes your current work

* **PDF Export**
  * Builds an A4 report with a cover, contents page, overall summary and metrics, each chart and pivot table with its insight, and a final recommendations section
  * Uses your browser's print dialog. Choose **Save as PDF**

* **User Authentication**
  * Sign up, log in and log out
  * All analysis pages are login-protected, and each user works in their own session

## How It Works

1. **Upload** a data file, or import a saved `.cascade` progress file.
2. **Clean** the data and decide how to handle rows with missing values.
3. **Visualize** by adding metrics, filters, pivot tables and charts, then edit the written parts.
4. **Export** the report as a PDF, and save your progress if you want to continue later.

## Limits

| Item | Limit |
| ---- | ----- |
| Upload size | 5 MB |
| Progress file size | 15 MB |
| Manual missing-value review | 500 affected rows |
| Near-duplicate check | Datasets up to 1,000 rows |
| Filters / Key Metric cards | 12 each |
| Top N / Bottom N | 1 to 50 |
| Pivot table display | 200 rows, 15 columns when a column field is used |
| Restored from progress file | 50 charts, 20 pivot tables, 500 columns |

## Tech Stack

| Technology | Purpose |
| ---------- | ------- |
| Python | Programming language |
| Django | Web framework, authentication and sessions |
| Pandas | Data cleaning, filtering, pivot tables and aggregation |
| NumPy | Numerical computation |
| Matplotlib | Chart rendering |
| openpyxl / xlrd | Reading `.xlsx` and `.xls` files |
| HTML/CSS/JavaScript | Front-end interface |
| SQLite | Development database |

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/YOUR-USERNAME/YOUR-REPOSITORY.git
cd YOUR-REPOSITORY
```

### 2. Create a virtual environment

```bash
python -m venv venv
```

### 3. Activate the virtual environment

**Windows:**

```powershell
venv\Scripts\activate
```

**macOS / Linux:**

```bash
source venv/bin/activate
```

### 4. Install dependencies

```bash
pip install -r requirements.txt
```

If a `requirements.txt` file is not available, install the main dependencies:

```bash
pip install django pandas numpy matplotlib openpyxl xlrd
```

### 5. Run migrations

```bash
python manage.py migrate
```

### 6. Start the development server

```bash
python manage.py runserver
```

Open the application at:

```text
http://127.0.0.1:8000/
```

## Usage

1. Sign up or log in.
2. Upload a CSV, Excel or JSON file, or import a saved `.cascade` progress file.
3. Press **Clean Data**, then choose how to handle any rows with missing values and confirm.
4. Review the cleaning report and continue to **Visualize**.
5. Set your filters and Key Metrics, then add the charts and pivot tables you need.
6. Review the generated insights, summary and recommendations, and edit them where needed.
7. Press **Save Progress** if you want to continue later.
8. Press **Download PDF** and choose **Save as PDF** in the print dialog.

Tip: finish your filters, metrics, charts and pivots first and edit the written text last. Changing those settings regenerates the related text.

## Documentation

A full guide to every page, button and limit is in the [Cascade User Manual](appdj/django_cascade_app/cascade_app/static/samples/Cascade_User_Manual.pdf).

## Data Storage

Uploaded and cleaned files are stored on the server that runs Cascade, and your charts, filters and notes live in your current login session. Use **Save Progress** to keep your work in a file. Only upload data you are allowed to store on that server.

## Purpose

Cascade was developed as a practical data analytics project focused on making data analysis more accessible.

Instead of requiring users to manually process datasets and create visualizations using separate tools, Cascade brings dataset cleaning, visualization, and insight generation into one web application.

> **Turn Raw Data Into Insight.**

## Development

This project is currently under development. Features and functionality may be improved or changed in future versions.

## Author

**Matt JLou D. Naparota**

BS Business Analytics
College of Business Administration
Silliman University

---

## License

This project is intended for educational and academic purposes.