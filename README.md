# Spearman Correlation Heatmap Generator

[![DOI](https://zenodo.org/badge/1013041179.svg)](https://doi.org/10.5281/zenodo.15797835)

Generates a heatmap of Spearman correlation coefficients with significant p-values annotated, from any Excel sheet.

---

## 📦 Usage

```bash
python spearman_heatmap.py "data.xlsx" "Sheet1" --title "Correlation Heatmap" --output "heatmap.png"

📋 Requirements
pandas

seaborn

matplotlib

numpy

scipy

openpyxl

Install all dependencies using:

bash
Copy
Edit
pip install -r requirements.txt
📄 Citation
If you use this tool in your research, please cite:

Rajpurohit, K. (2025). Spearman Correlation Heatmap Generator (v1.0.1) [Computer software]. Zenodo. https://doi.org/10.5281/zenodo.15797835

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
