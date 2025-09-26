# Retail Sales & Supply Chain Analysis

**Author:** Muhammet Berke Ağaya
**Student ID:** 5001230021

## Project Overview

This project provides an in-depth analysis of retail sales and supply chain data to extract actionable insights. Using **Python** for data preprocessing and analytics, and **Power BI** for interactive visualizations, the aim is to optimize sales, improve marketing strategies, and streamline supply chain management.

---

## Key Findings

### Overall Sales Performance

* Year-over-year growth from $0.49M (2014) to $0.75M (2017).
* Seasonal peaks in Q4 and slight dips in Q3, useful for forecasting.

### Regional Performance

* West and East regions lead in sales volume; California and New York are top contributors.
* Central region shows higher profit margins despite lower sales, suggesting efficient pricing strategies.

### Product Category Performance

* **Top revenue:** Technology.
* **High profit margins:** Office Supplies.
* **Underperforming:** Furniture (Furnishings, Tables) requires review.
* Key products contributing to overall sales: Binders (14.24%), Storage (9.7%), Phones (8.97%).

### Customer Insights

* Consumer segment drives 50.56% of sales ($1.16M).
* Corporate segment has fewer orders (30.74%, $0.71M) but higher average order size.
* Frequent buyers identified, e.g., ZC-21910 and WB-21850.

### Discounts & Returns

* High discounts increase sales but reduce profit, particularly for Tables.
* High return rates for Copiers (~10%), Machines (~10%), Phones (~15%) suggest quality or description issues.

### Sales Team Performance

* Top-performing representatives include Anna Andre and Chuck Magee.

---

## Recommendations

### Sales Strategy

* Focus on high-volume corporate clients and bundle popular products.
* Reward repeat buyers with loyalty programs.
* Limit deep discounts on low-margin items; promote high-margin products.

### Marketing Strategy

* Segment-targeted campaigns (B2B for Corporate, home office for Consumer).
* Cross-sell promotions and seasonal campaigns aligned with sales trends.
* Emphasize quality guarantees for high-return products.

### Supply Chain & Operations

* Investigate Phone product quality to reduce returns.
* Optimize inventory levels: reduce Bookcases stock, ensure Copiers safety stock.

---

## Tools & Technologies

* **Python:** Data preprocessing, analysis, and automation.
* **Power BI:** Interactive dashboards and business intelligence insights.
* **Excel:** Data validation, cleaning, and reporting.

---

## Automation Script

A Python script was used to detect and optionally highlight empty cells in Excel datasets:

```python
import openpyxl
from openpyxl.styles import PatternFill

def find_empty_cells(file_path, output_file, highlight=False):
    workbook = openpyxl.load_workbook(file_path)
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        empty_cells = []
        print(f"\nSheet: {sheet_name} - Empty Cells:")
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None or str(cell.value).strip() == '':
                    address = f"{cell.column_letter}{cell.row}"
                    empty_cells.append(address)
                    if highlight:
                        cell.fill = yellow_fill
        if empty_cells:
            print(f"Total {len(empty_cells)} empty cells found:")
            print(", ".join(empty_cells))
        else:
            print("No empty cells found in this sheet.")
    workbook.save(output_file)
    print(f"\nProcess completed. Results saved to '{output_file}'.")

input_file = r"C:\\Users\\berke\\OneDrive\\Masaüstü\\archive\\Retail-Supply-Chain-Sales-Dataset.xlsx"
output_file = r"C:\\Users\\berke\\OneDrive\\Masaüstü\\archive\\Retail-Supply-Chain-Sales-Dataset_analyzed.xlsx"

find_empty_cells(input_file, output_file, highlight=True)
```

---

## Conclusion

This project provides actionable insights for sales optimization, inventory management, and targeted marketing strategies. Regular monitoring via Power BI dashboards and cross-department collaboration is recommended to adapt to market changes effectively.
