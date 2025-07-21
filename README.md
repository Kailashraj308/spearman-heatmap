
# Spearman Correlation Heatmap Generator

[![DOI](https://zenodo.org/badge/1013041179.svg)](https://doi.org/10.5281/zenodo.15797835)

Interactive pythan script which generates a heatmap of Spearman correlation coefficients with significant p-values annotated, from any Excel sheet.


---

## 📦 Usage

```bash
python spearman_heatmap.py "data.xlsx" "Sheet1" --title "Correlation Heatmap" --output "heatmap.png"

# Spearman Correlation Heatmap Tool

This script generates a Spearman correlation matrix and a corresponding heatmap from data in an Excel file. It annotates significant correlations and can export results as CSV.

## **Requirements**

- Python 3.x
- Required Python packages:
  - `pandas`
  - `numpy`
  - `matplotlib`
  - `seaborn`
  - `scipy`
  - `openpyxl` (for reading `.xlsx` files)

Install them (if not already installed) with:
```bash
pip install pandas numpy matplotlib seaborn scipy openpyxl
```

## **How to Use**

1. **Prepare your Excel file**  
   - Ensure your data is in a single sheet.
   - The sheet should contain only numeric columns you wish to analyze.

2. **Run the script from the command line:**
```bash
python spearman_heatmap.py <EXCEL_FILE_PATH> <SHEET_NAME> [--title TITLE] [--output OUTPUT_FILE] [--export_csv]
```

### **Arguments**

- `<EXCEL_FILE_PATH>`: Path to your Excel file (e.g., `data.xlsx`)
- `<SHEET_NAME>`: Name of the sheet to analyze (e.g., `Sheet1`)
- `--title TITLE`: (Optional) Custom title for the heatmap.
- `--output OUTPUT_FILE`: (Optional) Output filename for the heatmap image (default: `spearman_correlation_heatmap.png`)
- `--export_csv`: (Optional) If provided, exports the correlation and p-value matrices as CSV files.

### **Example**

```bash
python spearman_heatmap.py data.xlsx Sheet1 --title "My Correlation Heatmap" --output my_heatmap.png --export_csv
```

This will:
- Read `data.xlsx` from the sheet named `Sheet1`
- Create a heatmap image saved as `my_heatmap.png` with the custom title
- Export `correlation_matrix.csv` and `pvalue_matrix.csv` for further analysis

## **Output**

- A heatmap image file with annotated Spearman correlation coefficients (and significance indicated by `*`)
- Optionally, two CSV files with the raw correlation and p-value matrices.

---

**Tip:**  
If you see errors about missing columns or empty data, check your Excel sheet to ensure it contains numeric data and no missing values in the columns you wish to analyze.

📄 Citation
If you use this tool in your research, please cite:

Rajpurohit, K., & Vibash Kalyaan, V. L. (2025). Spearman Correlation Heatmap Generator (Version 1.0.1) [Computer software]. https://doi.org/10.5281/zenodo.15797835

BibTeX:

bibtex
Copy
Edit
@software{rajpurohit2025heatmap,
  author       = {Kailash Rajpurohit},
  title        = {{Spearman Correlation Heatmap Generator}},
  version      = {v1.0.1},
  year         = 2025,
  publisher    = {Zenodo},
  doi          = {10.5281/zenodo.15797835},
  url          = {https://doi.org/10.5281/zenodo.15797835}
}
🔍 License
This project is licensed under the MIT License.
