# Bio-Composite Thesis Automation Tools

## 生物质复合材料毕业论文自动化工具集

A collection of Python automation tools built for the thesis **"Performance Evolution and Mechanism of KH570-Modified Nano-SiO₂/Bamboo Flour/PLA Composites Under Different Aging Conditions"**.

Built entirely with AI-assisted development (AI Agent + iterative code generation), demonstrating how large language models can accelerate scientific research workflows.

### 🎯 What Problem Does This Solve?

Academic researchers, especially in materials science, spend an enormous amount of time on repetitive tasks:

- Generating publication-quality charts with proper formatting (error bars, Chinese labels, black-and-white print compatibility)
- Processing experimental data across multiple aging conditions
- Managing Word document formatting (font specifications, mixed Chinese/English typesetting)
- Maintaining cross-references and citations
- Extracting, restructuring, and cleaning up thesis chapters

This toolkit automates all of the above, reducing what would take days of manual work to minutes of script execution.

### 📁 Repository Structure

```
thesis-automation-tools/
├── chart_generation/           # Scientific chart generation
│   ├── gen_v2_charts.py        # Mechanical properties bar charts with error bars
│   ├── gen_ftir_charts.py      # FTIR spectroscopy plots
│   ├── gen_stress_strain_curves.py  # Stress-strain curves
│   ├── gen_flexural_curves.py  # Flexural test curves
│   ├── make_aging_change_charts.py  # Aging degradation trend charts
│   ├── make_uv_vs_heat_comparison.py  # UV vs thermal aging comparison
│   ├── process_by_aging_condition.py  # Data grouping by aging condition
│   ├── process_by_aging_condition_bw.py  # B&W print variant
│   ├── process_time_course_data.py  # Time-course data processing
│   ├── gen_moisture_chart.py   # Moisture absorption charts
│   ├── gen_ch6_charts.py       # Chapter 6 charts
│   ├── gen_ch7_charts.py       # Chapter 7 charts
│   └── generate_thesis_charts.py  # Unified chart generation entry
│
├── doc_processing/             # Word document automation
│   ├── update_doc_v2.py        # Master document update (955 lines)
│   ├── modify_thesis.py        # Thesis content modification
│   ├── modify_thesis_chapters.py  # Chapter restructuring
│   ├── restructure_ch7.py      # Chapter 7 reorganization (689 lines)
│   ├── refresh_charts_in_doc.py   # Refresh chart images in DOCX
│   ├── refresh_all_charts.py   # Batch refresh all charts
│   ├── update_thesis_figures.py   # Update figure references
│   ├── extract_thesis.py       # Extract thesis text content
│   ├── extract_ch7.py          # Extract Chapter 7
│   ├── modify_doc_1_1.py       # Section 1.1 modifications
│   ├── change_to_wrap_top_bottom.py  # Table text wrapping
│   └── cleanup_residual_content.py   # Clean residual artifacts
│
├── reference_management/       # Citation & reference tools
│   ├── add_citations.py        # Add references from CrossRef
│   ├── fix_references_v2.py    # Fix reference formatting
│   └── fix_references_from_original.py  # Extract refs from source
│
├── requirements.txt            # Python dependencies
└── README.md                   # This file
```

### 🧠 Core Architecture & Logic Flow

The system follows a **data → visualization → document integration** pipeline:

1. **Data Layer**: Experimental data (mechanical testing, FTIR, moisture absorption) stored in structured formats
2. **Processing Layer**: Scripts aggregate, normalize, and group data by aging condition/type
3. **Visualization Layer**: Matplotlib generates publication-ready charts with:
   - Chinese axis labels with proper font rendering
   - Error bars on all bar charts
   - Pattern/texture fills for black-and-white print compatibility
   - Pre-computed statistical annotations
4. **Document Layer**: python-docx scripts inject charts into thesis chapters, update in-text references, fix mixed Chinese/English font issues (Song Ti for Chinese, Times New Roman for English/numbers)

### 🔬 Technical Highlights

- **Multi-agent logic flow**: Chart generation scripts call shared data-loading modules; document scripts operate on the same data pipeline — effectively a **single-agent orchestration** pattern where each script is a specialized worker
- **Long-chain reasoning**: The font-fixing scripts (e.g., modifying Heading 3 `ascii` font in styles.xml) required deep understanding of OOXML internals, XML traversal, and python-docx API intricacies
- **Print-optimized output**: All charts include a black-and-white mode (pattern fills, dashed line styles) for physical thesis printing

### ⚙️ Requirements

```bash
pip install -r requirements.txt
```

Scripts expect:
- Python 3.8+
- matplotlib, numpy, pandas, python-docx, lxml, openpyxl
- Experimental data in specific directory structure (adaptable)

### 🚀 Getting Started

1. Clone the repo
2. Place your experimental data in the expected paths (see individual script headers)
3. Run specific scripts or the master entry point

```bash
# Generate mechanical properties charts
python chart_generation/gen_v2_charts.py

# Generate FTIR spectra
python chart_generation/gen_ftir_charts.py

# Refresh all charts in thesis document
python doc_processing/refresh_charts_in_doc.py
```

### 📄 License

MIT

