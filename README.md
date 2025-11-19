# HR Data Analysis and Workforce Performance Insights (SIEHS)
**Author:** Muhammad Yahya  

---

## Confidentiality Notice
This repository contains no sensitive company data. All files are sanitized and authorized 

---

## Overview
This project analyzes workforce attendance and punctuality for **Sindh Integrated Emergency & Health Services (SIEHS)**.  
Multiple HR files were cleaned, standardized, and merged to create a unified dataset powering a Power BI workforce monitoring dashboard.

### Included in this repository
- Analysis Report (PDF)  
- R data-processing script  
- Master dataset generation pipeline  
- Dashboard preview image (PBIX not included)

---

## Objectives
- Integrate and harmonize five separate HR datasets  
- Convert Excel date/time formats into usable values  
- Clean invalid or inconsistent punch entries  
- Engineer KPIs such as:
  - Attendance and absences  
  - Working hours  
  - Late arrivals  
  - Leave breakdown  
  - Invalid entry counts  
- Build a unified master dataset  
- Produce workforce insights through analytics

---

## Method Summary

### 1. Data Cleaning & Standardization
- Normalized column names  
- Removed blank/duplicate entries  
- Converted Excel decimal dates & times  
- Standardized leave categories  
- Reconstructed Time In / Time Out pairs  

### 2. Attendance Transformation
- Reshaped daily attendance into long format  
- Computed daily and total working hours  
- Tagged invalid entries and absences  
- Identified off-days & holidays  

### 3. KPI Engineering
Metrics generated:
- Attended Days  
- Absent Days  
- Total Working Hours  
- Invalid Entries  
- Approved / Unapproved Leaves  
- Off Days & Holidays  
- Schedule status  

### 4. Master Dataset Construction
All cleaned datasets were merged using **Emp Index**, producing an analytics-ready employee-level dataset.

---

## Dashboard Insights
The dashboard provides:
- Attendance & punctuality trends  
- Department-level performance  
- Shift compliance analysis  
- Employee-level summaries  
- Workforce efficiency indicators  

*A static dashboard preview is included.*

---

## Contact
**Muhammad Yahya**  
Email: yahyaehtisham2004@gmail.com
