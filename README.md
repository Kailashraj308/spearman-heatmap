# Spearman Correlation Heatmap Generator

Generates a heatmap of Spearman correlation coefficients with significant p-values annotated, from any Excel sheet.

## Usage

```bash
python spearman_heatmap.py "data.xlsx" "Sheet1" --title "Correlation Heatmap" --output "heatmap.png"
```

## Requirements

- pandas
- seaborn
- matplotlib
- numpy
- scipy
- openpyxl

Install all dependencies using:

```bash
pip install -r requirements.txt
```
