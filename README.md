# Excel Salary Dashboard

An interactive Excel dashboard that shows **median salaries for data-related jobs**. You choose a job title, a country and a job type, and the dashboard returns the matching median salary. Two charts add the bigger picture: which roles pay more, and how salaries differ by country.

**Jump to:**
[Overview](#overview) ·
[How to use it](#how-to-use-the-dashboard) ·
[Dataset](#dataset) ·
[Excel skills used](#excel-skills-used) ·
[Dashboard build](#dashboard-build) ·
[Key takeaways](#key-takeaways) ·
[Conclusion](#conclusion)

---

## Overview

| Item | Detail |
|---|---|
| Project | Excel Salary Dashboard |
| Dashboard file | [`1_Salary_Dashboard.xlsx`](Region wise job analysis.xlsx) |
| Data | Real-world data science job information from 2023 (from my Excel course) |
| Main Excel skills | Charts, formulas and functions, data validation |
| Inputs the user controls | Job Title, Country, Type |
| Output | Median salary for the selected combination |

### Questions the dashboard helps answer

| Question | Where to find the answer |
|---|---|
| How do salaries compare across data-related job titles? | Bar chart (Chart 1) |
| How do salaries differ from country to country? | Map chart (Chart 2) |
| What is the median salary for a specific title, country and job type? | Median salary lookup (Formula 1) with the dropdowns |
| How do location and job type influence salaries? | Change the Country and Type dropdowns and compare results |

---

## How to use the dashboard

| Step | Action |
|---|---|
| 1 | Download and open [`1_Salary_Dashboard.xlsx`](1_Salary_Dashboard.xlsx). |
| 2 | Pick a value from each dropdown: **Job Title**, **Country** and **Type**. |
| 3 | Read the median salary returned for your selection. |
| 4 | Change one dropdown at a time to see how that factor affects the result. |

> **Note:** The dashboard uses `FILTER()` and an Excel map chart, so a recent version of Excel (Microsoft 365 recommended) works best.

---

## Dataset

The dataset contains real-world data science job information from 2023.

| Field | What it describes |
|---|---|
| Job titles | The role being advertised |
| Salaries | What the role pays |
| Locations | Where the job is based |
| Skills | What the employer asks for |

---

## Excel skills used

| Skill | How it is used in this dashboard |
|---|---|
| Charts | Horizontal bar chart (median salary by job title) and map chart (median salary by country) |
| Formulas and functions | A multi-criteria `MEDIAN(IF(...))` formula and a `FILTER()` formula that builds a clean list of job types |
| Data validation | Dropdown lists for Job Title, Country and Type that restrict input to valid options |

---

## Dashboard build

### Chart 1: Salaries by job title (bar chart)

| Aspect | Detail |
|---|---|
| Excel feature | Bar chart with formatted salary values and a layout optimized for clarity |
| Design choice | Horizontal bars, which make median salaries easy to compare |
| Data organization | Job titles sorted by descending salary for readability |
| Insight | Senior roles and Engineers are higher paying than Analyst roles |

### Chart 2: Median salary by country (map chart)

| Aspect | Detail |
|---|---|
| Excel feature | Excel's map chart, plotting median salaries globally |
| Design choice | Color-coded map that separates salary levels across regions |
| Data represented | Median salary for each country with available data |
| Visual benefit | Geographic salary trends are readable at a glance |
| Insight | Global salary differences are easy to see, including high and low salary regions |

### Formula 1: Median salary by job title, country and type

This formula returns the median salary for the job title, country and type the user selects.

```
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
```

**In plain English:** keep only the job postings that match the chosen title, country and job type and that have a real salary, then take the median of their yearly average salaries.

| Part of the formula | What it does |
|---|---|
| `jobs[job_title_short]=A2` | Matches the selected job title |
| `jobs[job_country]=country` | Matches the selected country |
| `ISNUMBER(SEARCH(type,jobs[job_schedule_type]))` | Checks that the selected type appears in the posting's schedule type text |
| `jobs[salary_year_avg]<>0` | Excludes blank or zero salaries |
| `IF(..., jobs[salary_year_avg])` | Returns the yearly average salary for postings that pass every check |
| `MEDIAN(...)` | Calculates the median of the returned salaries |

| Property | Detail |
|---|---|
| Type | Array formula using `MEDIAN()` with a nested `IF()` |
| Filters applied | Job title, country, schedule type, and non-blank salary |
| Output | The table that feeds the dashboard, giving the median salary for the specified job title, country and type |

### Formula 2: Clean list of job schedule types

This formula creates the list of job types used in the Type dropdown.

```
=FILTER(J2#,(NOT(ISNUMBER(SEARCH("and",J2#))+ISNUMBER(SEARCH(",",J2#))))*(J2#<>0))
```

**In plain English:** take the list of schedule types, remove any entry containing "and" or a comma, remove zeros, and keep the rest.

| Part of the formula | What it does |
|---|---|
| `J2#` | The spilled list of job schedule types |
| `ISNUMBER(SEARCH("and",J2#))` | Flags entries that contain "and" |
| `ISNUMBER(SEARCH(",",J2#))` | Flags entries that contain a comma |
| `NOT(...)` | Keeps only entries that were not flagged |
| `(J2#<>0)` | Removes zero values |
| `FILTER(...)` | Returns the remaining entries as the final list |

| Property | Detail |
|---|---|
| Purpose | Produces a list of unique job schedule types |
| Used for | The Type dropdown (see data validation below) |

### Data validation: dropdown lists

The filtered lists are applied as data validation rules (Data tab) to the **Job Title**, **Country** and **Type** options.

| Benefit | Explanation |
|---|---|
| Restricted input | Users can only choose from predefined, validated options |
| Fewer errors | Incorrect or inconsistent entries are prevented |
| Better usability | The dashboard is easier and safer to use |

---

## Key takeaways

| What was examined | What it showed |
|---|---|
| Job titles (bar chart) | Senior roles and Engineers are higher paying than Analyst roles |
| Countries (map chart) | Clear global salary differences, with high and low salary regions easy to identify |
| Dashboard design | Dropdowns backed by validated lists prevent incorrect entries and improve usability |

---

## Conclusion

I created this dashboard to show salary trends across various data-related job titles. Using data from my Excel course, the dashboard lets users make informed decisions about their career paths and explore how location and job type influence salaries.
