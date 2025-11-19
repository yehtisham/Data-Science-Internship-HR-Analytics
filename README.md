HR Data Analysis and Workforce Performance Insights (SIEHS)

Author: Muhammad Yahya

Confidentiality Notice:

This repository contains no sensitive company information.

Overview

This project analyzes attendance, punctuality, shift compliance, and workforce efficiency for Sindh Integrated Emergency & Health Services (SIEHS) using HR timestamp data from ~1,800 employees. The goal was to transform fragmented Excel sheets into a unified, analytics-ready master dataset and build an interactive Power BI dashboard enabling managers to monitor operational performance and identify workforce patterns.

This repository includes:

A professionally written Analysis Report (PDF)

Clean R code used for data ingestion, cleaning, transformation, and metric computation

The master dataset generation pipeline

A high-level view of the analytics dashboard (image only — not the PBIX file)

All data used is authorized for portfolio use and contains no confidential values.

Objectives

Integrate five independent HR data sources (attendance, invalid entries, leaves, absent sheets, pending leaves).

Correct and convert Excel date/time formats into meaningful metrics.

Detect and clean invalid, missing, or duplicated time entries.

Compute workforce KPIs, including:

Attended days

Absences

Approved vs. unapproved leaves

Working hours

Late arrivals

Invalid punch-ins

Build a unified master dataset for analytics.

Generate visual insights for senior management.

Data Sources (Raw Inputs)

Methodology Summary
1. Data Cleaning & Standardization

Key processing steps:

Normalized column names with janitor::clean_names()

Removed duplicated or blank rows

Converted Excel decimal dates using custom functions

Reconstructed time-in / time-out pairs

Standardized department and position labels

Handled multi-format leave types (annual, casual, sick, compensatory, etc.)

2. Attendance Transformation

The attendance sheet was reshaped using pivot_longer() into a long format to:

Compute daily working hours

Detect invalid entries

Flag absences, off-days, and holidays

Aggregate totals per employee

3. KPI Engineering

Metrics computed:

Total Assigned Duties

Total Attended Days

Total Absences

Working Hours

Invalid Entries

Approved / Unapproved Leaves

Off Days & Holidays

Schedule Status

Leave Type Breakdown

4. Master Dataset Construction

All datasets were merged on Emp Index, producing a consolidated table with one row per employee, covering onboarding data, attendance, behavior, and leave patterns.

The final dataset powers the Power BI dashboard.

Dashboard Insights

The Power BI dashboard highlights:

On-time arrival trends

Shift-level attendance compliance

Department-level punctuality performance

Late arrivals heatmaps

Employee-level summaries

Workforce efficiency indicators

A static dashboard image is included for reference.

Contact

For questions or verification:

Muhammad Yahya
Email: yahyaehtisham2004@gmail.com
