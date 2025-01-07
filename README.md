# VBA Stock Analysis

## Overview
This project demonstrates the development of a VBA script to perform stock market analysis as part of the Week 2 homework assignment. The script efficiently processes data to summarize key metrics for each stock ticker symbol in a given dataset.

---

## Features
The VBA script calculates the following metrics for each ticker symbol:
1. **Yearly Change**: The difference between the opening price at the beginning of the year and the closing price at the end of the year.
2. **Percentage Change**: The percentage change from the opening price at the beginning of the year to the closing price at the end of the year.
3. **Total Stock Volume**: The cumulative stock volume for the ticker over the year.

---

## Deliverables
- **VBA Script**: The VBA code used to perform the analysis.
- **Analysis Results**: Screenshots showcasing the results of the analysis for each year.

---

## How It Works
1. **Data Iteration**:
   - The script loops through the dataset to identify changes in ticker symbols.
   - It captures the opening price at the beginning of the year and the closing price at the end of the year.
2. **Metric Calculation**:
   - Calculates the yearly change and formats it to visually indicate positive or negative changes (e.g., green for positive, red for negative).
   - Computes the percentage change, handling division by zero where applicable.
   - Aggregates the total stock volume for each ticker.
3. **Results Output**:
   - Outputs the calculated metrics into a summary table for easy review.

---

## Screenshots
Screenshots of the analysis results for each year are provided in the repository to validate the script's output and functionality.

---

## Repository Contents
- `VBA_Stock_Analysis.bas`: The VBA script file.
- `Screenshots/`: Folder containing screenshots of the analysis results for each year.
- `README.md`: Project documentation.

---

## Getting Started
To replicate the analysis:
1. Open the dataset in Microsoft Excel.
2. Load the provided VBA script into the Visual Basic for Applications editor.
3. Run the script and observe the generated summary table.

---

## Notes
- Ensure that the dataset is properly formatted before running the script.
- The script is designed to handle large datasets efficiently but may require adjustments for different data structures.

---

## Acknowledgments
This project was completed as part of a learning module to develop proficiency in VBA scripting for data analysis tasks.

