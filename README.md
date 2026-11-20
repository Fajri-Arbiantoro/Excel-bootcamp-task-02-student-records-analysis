# Project 02: Student Records Analysis Using Intermediate Excel

My second Excel Bootcamp project, documenting my progress in using Excel to clean, connect, process, and analyse a student-records practice dataset.

## Project Background

This project was completed as an intermediate post-test and personal task during an Excel Bootcamp I participated in.

The assignment involved working with 750 student records across multiple worksheets. The objective was to connect related datasets, handle missing attendance records, apply scoring rules, calculate final grades, and summarize the results by major.

Compared with my first Excel project, which focused on cleaning sales records, this assignment gave me more practice combining lookup functions, logical formulas, conditional calculations, and summary analysis across linked worksheets.

This project represents another step in developing practical data analysis skills using Microsoft Excel.

## Project Files

- [Anonymized Excel Workbook](./Student-Records-Analysis-Anonymized-CHECKED.xlsx)
- [Project Presentation](./Excel%20Portfolio%20Part%202%20Fajri%20Arbiantoro.pdf)

## Learning Objectives

Through this assignment, I aimed to:

- Generate consistent student IDs using text and logical functions.
- Connect information across multiple worksheets.
- Retrieve data using VLOOKUP and INDEX + MATCH.
- Handle missing attendance records with IFERROR.
- Apply conditional rules to assignment scores.
- Calculate weighted final scores and assign grades.
- Summarize records using conditional counting and aggregation.
- Use conditional formatting and charts to make results easier to interpret.

## Tasks & Analysis Process

**1. Generating Student IDs**

Created a consistent student ID by combining a major code, zero padding, and the student number.

Functions used: LEN, IFS, CONCAT.

**2. Linking Worksheets**

Used student IDs to retrieve names and other fields from the Detail worksheet.

Functions used: VLOOKUP, INDEX, MATCH.

**3. Handling Missing Attendance Records**

Used IFERROR to return "No" when no matching attendance record was found. An absence-date rule deducted 10 points from the assignment score.

Functions used: IFERROR, ISNUMBER, IFS.

**4. Calculating Final Scores and Grades**

Calculated final scores using six weighted components:

- Quiz 1: 15%
- Quiz 2: 15%
- UTS: 20%
- Quiz 3: 15%
- UAS: 20%
- Updated assignment score: 15%

The final score was classified into grades from A+ to E.

Functions used: IFS, AND, and weighted calculations.

**5. Conditional Formatting and Data Checks**

Applied conditional formatting to highlight:
- Green row: final score above 80 with no absence record.
- Yellow cell: no absence record.
- Red cell: grade C.
- Pink cells: duplicate values flagged as a data-quality check.

COUNTIF and COUNTIFS were used to cross-check counts by grade and attendance status.

**6. Summary Tables and Charts**

Created summary tables and charts to compare results across four majors, including student distribution, quiz scores, attendance records, and average height and weight.

Functions practised included COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, SUMIFS, ROUNDUP, MAX, MIN, INDEX, MATCH, YEARFRAC, and TODAY.

## Key Analysis

Based on the summary tables and charts:

1. Student distribution was relatively balanced across the four majors, ranging from 23% to 26%.
2. Average Quiz 1 and Quiz 2 scores differed by approximately two points across majors.
3. Additional summaries compared Quiz 3 totals and attendance records by major.

## Tools & Techniques

**Tool:** Microsoft Excel

**Skills Practised:**
- Data Preparation
- Student ID Generation
- Worksheet Linking
- Lookup Functions
- Error Handling
- Logical Functions
- Weighted Calculations
- Grade Classification
- Conditional Formatting
- Duplicate Detection
- Conditional Aggregation
- Summary Tables
- Data Visualization

## Challenges & Learning Reflection

One of the main challenges in this project was understanding how multiple formulas could work together across different worksheets. Generating consistent IDs made it possible to connect the records, while INDEX + MATCH helped retrieve information that VLOOKUP could not return from the left side of the lookup column.

I also learned how to handle missing attendance records without allowing lookup errors to interrupt the calculations. Combining attendance rules with weighted scoring helped me understand how academic rules can be translated into spreadsheet logic.

Another important takeaway was the value of checking results. Conditional formatting, COUNTIF, and COUNTIFS provided additional ways to review the data and compare displayed results with summary calculations.

## What's Next?

This project builds on the data-cleaning techniques I practised in my first Excel assignment.

I plan to continue strengthening my Excel skills and gradually expand my toolkit with SQL, Python, and data visualization tools as I work toward a career in data analytics.

I will continue documenting each project, including the techniques I practise, the challenges I encounter, and how my analytical skills develop over time.

---

**Disclaimer**

This project was completed as part of an intermediate Excel Bootcamp for educational and portfolio purposes only. The analysis is based on a student-records practice dataset supplied for the assignment. The workbook shared in this repository is an anonymized version prepared for portfolio demonstration. The original, non-anonymized workbook is not shared because it may contain personal information about students. Please do not attempt to identify individuals or use the dataset for purposes beyond this educational demonstration. The project is not intended for commercial use or to represent actual student performance.
