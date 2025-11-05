# Excel Comparison Tool

A React-based web application to compare two Excel files side-by-side with visual highlighting of differences.

## Features

- 📊 Upload and compare two Excel files (.xlsx, .xls)
- 🔍 Side-by-side comparison with color highlighting
- ✅ Green cells = Matching data
- ❌ Red cells = Different data
- 📑 Multiple sheet support
- 📥 Export comparison results to Excel
- 📈 Summary statistics (total cells, matches, differences)

## Installation
```bash
# Clone the repository
git clone https://github.com/z-onu/excel-comparison-tool.git

# Navigate to project directory
cd excel-comparison-tool

# Install dependencies
npm install

# Start development server
npm start
```

## Usage

1. Click "Upload First Excel File" and select your first file
2. Click "Upload Second Excel File" and select your second file
3. Select which sheets to compare (if multiple sheets exist)
4. View the comparison results with color-coded highlighting
5. Click "Export to Excel" to download the comparison report

## Technologies Used

- React 18
- SheetJS (xlsx) - Excel file processing
- Lucide React - Icons
- CSS3 - Styling

## Deployment

To deploy to GitHub Pages:
```bash
npm run deploy
```

## Licensed