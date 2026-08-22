# Excel Salary Dashboard

An interactive Excel dashboard for exploring data-job compensation by role, country, and work arrangement. The project demonstrates practical Excel dashboard design, formula-driven analysis, and user-friendly filtering.

![Excel Salary Dashboard](https://github.com/user-attachments/assets/230e121b-bc4b-4334-bd18-293cec7d685d)

## Business Objective

The dashboard helps users compare salary levels across data-related roles and locations so they can evaluate career opportunities and compensation benchmarks.

## Tools & Techniques

- Microsoft Excel
- Advanced formulas and functions
- Dynamic arrays
- Data Validation
- Charts and dashboard design
- Geographic Map Charts

## Dashboard Features

### Salary by Job Title

A horizontal bar chart ranks median salaries by job title, making differences between analyst and senior technical roles easy to compare.

![Salary by Job Title](https://github.com/user-attachments/assets/286a6b11-e99d-466e-808f-6db368d465ef)

### Salary by Country

The map visualization shows median salary by country and provides a geographic view of compensation differences.

![Country Salary Map](/0_Resources/Images/1_Salary_Dashboard_Country_Map.gif)

### Interactive Filters

Data Validation controls allow users to select job title, country, and schedule type from predefined lists. This makes the dashboard easier to use and reduces inconsistent inputs.

![Data Validation](https://github.com/user-attachments/assets/ccea0c6d-0faa-4b03-a2a4-301fcab28e52)

## Example Formula

The dashboard uses multi-criteria calculations to return median salary for a selected job title, country, and schedule type:

```excel
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

A dynamic `FILTER()` formula is also used to generate valid schedule-type selections.

## Key Takeaways

- Senior and specialized data roles generally show higher median salaries than analyst roles.
- Compensation varies substantially by country.
- Interactive filtering makes the dashboard useful for comparing different career scenarios.

## Portfolio Value

This project demonstrates my ability to combine **Excel formulas, data validation, visualization, and dashboard design** to answer practical business questions and communicate results clearly.
