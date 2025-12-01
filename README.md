** Excel_Analysis**

**Project Overview**
This Excel capstone project analyzes healthcare data containing medical examinations, hospitalization details, and customer profiles. The goal is to clean, transform, and merge datasets to uncover insights related to patient profiles, medical history, healthcare utilization, and total charges. The final output includes an interactive dashboard with slicers and visual analytics.

**1. Data Cleaning
** Missing Values**

* Counted missing values (?) in Medical Examinations and Hospitalization Details tables.
* Filled missing month with Sep and year with the rounded average year.
* Treated missing smoker, Hospital tier, and City tier using the mode.
* Replaced missing State ID with “Unknown”.

**2. Data Transformation**

* Split names column into Title, First Name, and Last Name.
* Converted NumberOfMajorSurgeries to numeric by standardizing inconsistent entries.
* Checked for inconsistencies in Heart Issues and smoker and corrected them.
* Created Weight Status using BMI categories.
* Created Diabetes Status using HbA1C ranges.
* Merged year + month + date into a Date of Birth column.
* Calculated Age using DATEDIF based on 08-Jun-2023.
* Formatted charges column as $ currency.

 **3.Data Exploration, Analysis & Visualization**
 
**Created a new sheet “Healthcare”, merging all tables using Customer ID via VLOOKUP.**
Retained:
Customer ID, First Name, BMI, HbA1C, Heart Issues, Any Transplants, Cancer history, NumberOfMajorSurgeries, smoker, Weight Status, Diabetes Status, Date of Birth, charges, Hospital tier, City tier, State ID, Age.

**4. Analysis and Charts**

**Pie/Donut Chart Analysis**
* Distribution of cancer history among smokers vs non-smokers.
* Total number of major surgeries & average HbA1C for transplant vs non-transplant patients.

**Column/Bar Chart Analysis**
* Charges comparison based on Weight Status and Diabetes Status.
* Average charges across Hospital Tiers for each State.

**Line/Scatter Plot Analysis**
* Correlation between Age vs BMI and Age vs HbA1C.
* Relationship between Age vs Healthcare Charges.

**5.Dashboard Development**

* Built an interactive Excel dashboard with:
* Donut, column, bar, line, and scatter charts
* Slicers for Weight Status and Diabetes Status
* Clean formatting for clear storytelling
* Unified view of medical risks, costs, and demographics

**Key Insights**
* Clear patterns between BMI, diabetes, age, and health risks
* Charges vary significantly across weight & diabetes categories
* Strong differences in cost across hospital tiers
* Smokers show different cancer history proportions
* Age correlates with higher BMI, HbA1C, and charges in certain groups

**Tools Used**
* Microsoft Excel
* Pivot Tables
* Advanced formulas (VLOOKUP, DATEDIF, IF functions)
* Data visualization techniques


  

