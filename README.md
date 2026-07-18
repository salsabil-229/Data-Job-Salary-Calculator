# Data Job Market Salary Calculator & Interactive Dashboard

## Project Overview
As an aspiring Data Analyst entering the competitive job market, I wanted to build a practical tool that answers three critical questions:
1. **What is the typical salary for a specific data role in a given country?**
2. **How many active job listings exist for these criteria?**
3. **Where are these companies actively hiring (which job boards/platforms)?**

This project is an **interactive, formula-driven Excel Dashboard & Calculator** designed to help job seekers estimate their market value, prepare for salary negotiations, and identify the best platforms to target during their job search.

---

## Interactive Features & Dashboard Preview

![Interactive Salary Dashboard](data_job_salary_dashboard.png)

* **Dynamic Dropdowns:** Users can select a specific **Job Title** (e.g., Data Analyst, Data Scientist, Data Engineer) and **Country** (e.g., United States, United Kingdom, India).
* **Instant KPI Cards:** The dashboard instantly recalculates and displays the **Number of Jobs Available**, **Median Salary ($ USD)**, and the **Top Platform to Apply**.
* **Clean User Interface (UI):** Gridlines were removed, and structured card elements were styled using a clean corporate color palette to simulate a modern web application.

---

## Excel Skills Demonstrated
* **Table Creation:** Creating a table of the **Data** sheet to improve efficiency and simplify formulas.
* **Dynamic Range Management:** `FILTER`, `UNIQUE`, and helper columns to construct clean, comma-free validation lists.
* **Advanced Formulas:** Multi-criteria `COUNTIFS`, custom nested array formulas (`MEDIAN` combined with `IF`), and `XLOOKUP`.
* **Data Validation:** Dropdown menus linked dynamically to automated background tables.
* **UI/UX Design:** Canvas formatting, gridline removal, strategic typography, and card-based layout structure.

---

## How the Calculator Engine Was Built 
Unlike standard dashboards that rely on Pivot Tables, this project is powered entirely by **dynamic background cell linking and logical formulas** divided across three distinct tabs: `Data` (raw data), `Data_Validation`, and `Calculator` (user interface).

### 1. Dynamic Dropdown Lists
To populate the dropdown menus without manual upkeep, I used the `UNIQUE` and `FILTER` functions on the **`Data_Validation`** tab to extract distinct Job Titles, Countries and Job Types from the raw dataset. 

* **Data Cleaning Logic:** To ensure messy, comma-separated job profiles didn't clutter the dropdowns, I filtered out strings containing commas and any conjunctions via a two-step process:
  ```excel
  =UNIQUE(jobs[job_schedule_type])
  ```
  ```excel
  =FILTER(L2#,NOT(ISNUMBER(SEARCH("and",L2#)))*(L2#<>0))
  ```

### 2. Live Job Counter (COUNTIFS)
To dynamically calculate the total volume of job postings matching the user's filtered criteria:  
  ```excel
  =COUNTIFS(jobs[job_country], Country, jobs[job_title_short], A2, jobs[job_schedule_type], "*" & Type & "*", jobs[salary_year_avg], ">0")
```
* **How it works:** The formula maps the exact criteria selected in the calculator/dashboard. The `"*" & Type & "*"` ensures that any title that includes the Job Type is counted and not only titles that are strictly equal to the Job Type. The `">0"` allows us to make sure that only positive yearly salary figures are taken into account, excluding and blanks or text errors.

### 3. Multi-Criteria Median Salary Estimator (MEDIAN + IF)
Because Excel does not have a native MEDIANIFS function, I designed a nested array formula to isolate the median salary of jobs that simultaneously meet both filter conditions:
```excel
=MEDIAN(IF((jobs[job_country]=A2)*(jobs[salary_year_avg]<>0)*(jobs[job_title_short]=Title)*(ISNUMBER(SEARCH(Type,jobs[job_schedule_type]))),jobs[salary_year_avg]))
```
* **How it works:** This formula acts as a logical filter, instructing Excel to evaluate the MEDIAN of the salary column only for the rows where all criteria evaluate to TRUE.

### 4. Data Retrieval (XLOOKUP)
Rather than running heavy calculations on the front-end dashboard, the `Calculator` tab acts as a lightweight display screen. It uses `XLOOKUP` to instantly pull the pre-calculated metrics from the background tabs:
* **Retrieving Live Job Counts (from Data_Validation tab):**
```excel
=XLOOKUP(Title,D2:D11,E2:E11,"No Result")
```
* **Retrieving Median Salary (from Title tab):**
```excel
=XLOOKUP(Title,E2:E11,F2:F11, "No Result")
```

### 5. Top Job Platform Processing (Sorting & Text Cleaning)
To determine the best platform for a job seeker to apply on, I avoided complex front-end lookup arrays and instead built a clean data pipeline on the `Platform` tab:
* **Count & Sort:** The total jobs for each platform were calculated using `COUNTIFS` and then sorted in descending order, automatically pushing the platform with the highest volume to the very top row.

* **Text Cleaning:** Because the raw data outputs platforms with a prefix (e.g., "via LinkedIn"), text manipulation (substitution) was used on the top result to strip out the word "via ". This ensures only the clean, professional brand name (e.g., "LinkedIn") is fed to the dashboard's KPI card.
