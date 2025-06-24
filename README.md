# EXCEL--Call-Center-Report
This project showcases an interactive call center dashboard developed using Excel tools such as Pivot Tables, Power Query, VLOOKUP, Data Validation, VBA, and AI automation for generating customer certificates. The goal is to analyze call center performance and customer satisfaction to drive strategic insights.

📊 Project Overview
The dataset consists of customer call records and customer demographic data, combined and transformed to answer key business questions regarding call activity, satisfaction, and revenue contribution.

Tools & Technologies Used:

a. Microsoft Excel
- Pivot Tables
- Power Query for ETL
- VLOOKUP for data merging
- Charts & Conditional Formatting
- VBA for certificate automation
- AI tools for name parsing and batch processing

🧠 Business Understanding
A telecom company aims to understand:

- How customers interact with their support center
- Which representatives bring the most revenue
- Who their top customers are
- When the call center is most busy
- Whether satisfaction ratings correlate with call duration

🔄 Data Preparation (ETL)

- 	Power Query: Used to clean and transform raw tables.
- 	VLOOKUP: Merged customer demographic data into the call data using Customer ID.
- 	Calculated Columns:
   Financial Year from Date of Call, Day of Week from Date of Call, Duration List (Short, Medium, Long),  Rating Rounded for bar charts

📈 Dashboard Features
Key Metrics:

-	Total Calls: 1,000
-	Total Amount: $96,623
-	Total Duration: 89,850 sec
-	Avg. Rating: 3.9
-	Happy Customers (Rating ≥ 4): 307

🔍 Analytical Questions Answered
✅ 1. How many calls are we getting by customer?
- Pivot table by Customer ID
- Sorted to identify frequent callers

✅ 2. How satisfied are our customers?
- Average rating: 3.9
- Rating distribution chart (1 to 5 scale)

✅ 3. Who are our top 10 customers?
- Based on total purchase amount
- Shown by sorting the pivot on purchase

✅ 4. Top 10 customers for a specific representative?
- Filter applied on pivot using Representative
- Sorted by amount

✅ 5. Call duration analysis
- Duration bucketed (short/med/long)
- Representative-wise duration pivot

✅ 6. How busy is our call centre in 2023?
- Monthly call trends chart (Jan–Dec)
- High in March & October, low in July

✅ 7. YTD Sales Analysis
- Total purchase amount YTD: $96,623

✅ 8. Which days of the week are the busiest?
- Saturday has the highest call count: 161
- Weekday-wise distribution shown in bar

✅ 9. Is there a link between call duration and satisfaction rating?
- Scatter chart and pivot explored
- Medium-duration calls show higher average ratings

✅ 10. Should we hire extra people in specific months?
- March and October show call spikes
- Recommend staffing based on trend

