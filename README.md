# PathwayColorMapper: PowerPoint-Based Visualization of Gene Expression Data on Pathway Diagrams

**PathwayColorMapper** is a Python tool that allows researchers to visualize gene-associated data, such as gene expression levels, directly on pathway diagrams within PowerPoint files. This enables intuitive, presentation-ready mapping of complex datasets using customizable color scales.

## Features
- **Custom Pathway Diagram Support**: Annotate any PowerPoint-based pathway diagram with your data.
- **Color Mapping**: Use gradient color scales to represent numerical data such as upregulation and downregulation.
- **Automated Processing**: Automatically map data values to diagram regions.
- **PowerPoint Output**: Create editable PowerPoint slides for further customization and presentations.
- **Flexible Data Input**: Accepts standard data formats like CSV or Excel for easy integration.

---

## Project Structure

```
GenePathwayPlot/
├── README.md                     # Project overview and usage instructions
├── LICENSE                       # MIT License file
├── config/
│   └── config.yaml               # Configuration file for color scales and settings
├── examples/
│   ├── example_data.csv          # Example input gene expression data (CSV)
│   ├── example_pathway.pptx      # Example PowerPoint pathway diagram
│   ├── example_output.png        # Example output image
├── output/                       # Output directory for generated files
│   ├── colorbar_vertical.png     # Generated vertical colorbar
│   ├── colorbar_horizontal.png   # Generated horizontal colorbar
│   ├── example_output.pptx       # Example annotated PowerPoint file
├── scripts/
│   └── gene_pathway_plot.py      # Main script for pathway annotation
├── requirements.txt              # Python dependencies
├── .gitignore                    # Files and directories to ignore in version control
```

---

## Installation

### Requirements
- Python 3.8+
- Libraries listed in `requirements.txt`.

### Steps
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/PathwayColorMapper.git
   ```
2. Navigate to the directory:
   ```bash
   cd PathwayColorMapper
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

### Input Data Format
Prepare your data in a CSV or Excel file with the following format:

| Gene | Value  |
|------|--------|
| TP53 | 2.5    |
| EGFR | -1.2   |
| BRCA1 | 1.8    |
| MYC  | 0.5    |
| PIK3CA | -0.8  |


### Steps to Generate Pathway Visualization
1. Prepare your input files:
   - Gene expression data in CSV or Excel format (see `examples/example_data_scaled.csv`).
   - Pathway diagram in PowerPoint format (see `examples/example_pathway.pptx`).
2. Run the script:
   ```bash
   python scripts/pathway_color_mapper.py --input data.csv --pathway pathway.pptx --output output.pptx
   ```
3. Open the output file:
   - `output.pptx` contains the annotated pathway with color-coded genes.

---

## Example

### Input Pathway Diagram
A blank or pre-designed pathway diagram in PowerPoint:

![Example Input](examples/example_input.png)

### Output Pathway Diagram
Annotated pathway with a heatmap overlay:

![Example Output](examples/example_output_scaled.png)

---

## Example Data
Sample input files are provided in the `examples/` directory to help you get started.

### Files
- `examples/example_data_scaled.csv`: A CSV file containing sample gene expression data.
- `examples/example_data_scaled.xlsx`: An Excel version of the same data.
- `examples/example_pathway.pptx`: A blank PowerPoint pathway diagram.

### How to Use
Run the script with the provided sample data:
```bash
python scripts/pathway_color_mapper.py --input examples/example_data_scaled.csv --pathway examples/example_pathway.pptx --output output/example_output_scaled.pptx
```
This will generate an annotated pathway diagram in `output/example_output_scaled.pptx`.

---

## Configuration

The behavior of the script is controlled by `config/config.yaml`. Customize the following settings:

- **Color Scale**: Define the gradient and range for color mapping.
- **Input/Output Settings**: Adjust file paths and additional parameters.

Example `config.yaml` for Z-score normalized data:
```yaml
color_scale:
  min_value: -2.0  # Minimum value for the color scale
  max_value: 2.0   # Maximum value for the color scale
  gradient:        # Colors for the color scale
    - color: blue  # Corresponds to negative values
    - color: white # Corresponds to mid values
    - color: red   # Corresponds to positive values
```

Example `config.yaml` for raw data (e.g., TPM):
```yaml
color_scale:
  min_value: 0       # Minimum value for the color scale
  max_value: 2.0     # Maximum value for the color scale
  gradient:  viridis # Colors for the color scale
```

Mapped pathway with example raw data:

![Example Output](examples/example_output.png)

---

## Dependencies
- `matplotlib`
- `python-pptx`
- `pandas`
- `openpyxl`
- `numpy`
- `pyyaml`

Install them using:
```bash
pip install -r requirements.txt
```

---

## Contributing
We welcome contributions! If you have suggestions for improvements, encounter bugs, or want to add features, feel free to submit an issue or pull request.

---

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Acknowledgements
PathwayColorMapper was created to simplify the visualization of complex genomic datasets in a pathway context. Special thanks to the open-source community for making tools like `python-pptx` and `matplotlib` available for projects like this.
