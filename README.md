# Cascade

### Turn Raw Data Into Insight

Cascade is a Django-based data analysis and visualization web application designed to help users transform raw datasets into clean, meaningful, and visual insights.

It allows users to upload datasets, clean and analyze their data, generate visualizations, and produce insights and recommendations from the results.

## Features

* **Dataset Upload**

  * Supports CSV, XLS, and XLSX files
  * Upload raw datasets directly through the application

* **Data Cleaning**

  * Review and clean uploaded datasets
  * Identify and handle common data issues before analysis

* **Data Visualization**

  * Automatically generate charts based on the selected data
  * Supports:

    * Bar charts
    * Horizontal bar charts
    * Pie charts
    * Line charts
    * Area charts
    * Scatter plots
    * Histograms
    * Box plots
    * Heatmaps

* **Data Insights**

  * Generate analytical insights based on the selected variables
  * Provides summaries and recommendations to help interpret the results

* **PDF Export**

  * Export analysis results and visualizations into a PDF report

* **User Authentication**

  * Login-protected analysis features
  * User-specific analysis and reports

## Tech Stack

| Technology | Purpose                      |
| ---------- | ---------------------------- |
| Python     | Programming language         |
| Django     | Web framework                |
| Pandas     | Data processing and analysis |
| NumPy      | Numerical computation        |
| Matplotlib | Data visualization           |
| HTML/CSS   | Front-end interface          |
| SQLite     | Development database         |

```

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

### 4. Install dependencies

```bash
pip install -r requirements.txt
```

If a `requirements.txt` file is not available, install the main dependencies:

```bash
pip install django pandas numpy matplotlib openpyxl
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

1. Log in to the application.
2. Upload a CSV, XLS, or XLSX dataset.
3. Review and clean the uploaded data.
4. Select the variables you want to analyze.
5. Generate a visualization.
6. Review the generated insights and recommendations.
7. Export the results as a PDF report.

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
