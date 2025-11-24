import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, AlertCircle, Check, Download } from 'lucide-react';
import './App.css';

export default function ExcelComparisonTool() {
  const [file1Data, setFile1Data] = useState(null);
  const [file2Data, setFile2Data] = useState(null);
  const [file1Name, setFile1Name] = useState('');
  const [file2Name, setFile2Name] = useState('');
  const [selectedSheet1, setSelectedSheet1] = useState('');
  const [selectedSheet2, setSelectedSheet2] = useState('');
  const [comparison, setComparison] = useState(null);
  const [error, setError] = useState('');

  const handleFile1Upload = async (e) => {
    const file = e.target.files[0];
    if (file) {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        const sheets = {};
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          sheets[sheetName] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        });

        setFile1Data({ sheets, sheetNames: workbook.SheetNames });
        setFile1Name(file.name);
        setSelectedSheet1(workbook.SheetNames[0]);
        setError('');
      } catch (err) {
        setError(`Error reading file 1: ${err.message}`);
      }
    }
  };

  const handleFile2Upload = async (e) => {
    const file = e.target.files[0];
    if (file) {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        const sheets = {};
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          sheets[sheetName] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        });

        setFile2Data({ sheets, sheetNames: workbook.SheetNames });
        setFile2Name(file.name);
        setSelectedSheet2(workbook.SheetNames[0]);
        setError('');
      } catch (err) {
        setError(`Error reading file 2: ${err.message}`);
      }
    }
  };

  React.useEffect(() => {
    if (file1Data && file2Data && selectedSheet1 && selectedSheet2) {
      const sheet1 = file1Data.sheets[selectedSheet1];
      const sheet2 = file2Data.sheets[selectedSheet2];

      const maxRows = Math.max(sheet1.length, sheet2.length);
      const maxCols = Math.max(
        ...sheet1.map(row => row.length),
        ...sheet2.map(row => row.length)
      );

      const diffs = [];
      for (let i = 0; i < maxRows; i++) {
        for (let j = 0; j < maxCols; j++) {
          const val1 = sheet1[i]?.[j] ?? '';
          const val2 = sheet2[i]?.[j] ?? '';
          if (val1 !== val2) {
            diffs.push({ row: i, col: j, val1, val2 });
          }
        }
      }

      setComparison({
        sheet1,
        sheet2,
        maxRows,
        maxCols,
        diffs,
        totalCells: maxRows * maxCols,
        matchingCells: (maxRows * maxCols) - diffs.length
      });
    }
  }, [file1Data, file2Data, selectedSheet1, selectedSheet2]);

  const getColumnLetter = (col) => {
    let letter = '';
    let num = col;
    while (num >= 0) {
      letter = String.fromCharCode(65 + (num % 26)) + letter;
      num = Math.floor(num / 26) - 1;
    }
    return letter;
  };

  const handleDownloadResults = () => {
    if (!comparison) return;

    const workbook = XLSX.utils.book_new();
    const usedSheetNames = new Set();

    const appendSheet = (sheetData, desiredName) => {
      const baseName = desiredName || 'Sheet';
      let uniqueName = baseName;
      let counter = 1;
      while (usedSheetNames.has(uniqueName)) {
        uniqueName = `${baseName}_${counter++}`;
      }
      usedSheetNames.add(uniqueName);
      XLSX.utils.book_append_sheet(workbook, sheetData, uniqueName);
    };

    const summaryData = [
      ['File 1', `${file1Name || 'File 1'} (${selectedSheet1 || 'Sheet'})`],
      ['File 2', `${file2Name || 'File 2'} (${selectedSheet2 || 'Sheet'})`],
      ['Total Cells', comparison.totalCells],
      ['Matching Cells', comparison.matchingCells],
      ['Differences', comparison.diffs.length]
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
    appendSheet(summarySheet, 'Summary');

    const diffRows = [
      ['Cell', 'File 1 Value', 'File 2 Value'],
      ...comparison.diffs.map(diff => [
        `${getColumnLetter(diff.col)}${diff.row + 1}`,
        diff.val1 === undefined ? '' : String(diff.val1),
        diff.val2 === undefined ? '' : String(diff.val2)
      ])
    ];
    if (diffRows.length === 1) {
      diffRows.push(['-', 'All cells match', '']);
    }
    const diffSheet = XLSX.utils.aoa_to_sheet(diffRows);
    appendSheet(diffSheet, 'Differences');

    const buildTable = (sheet) => {
      const head = ['Row/Col', ...Array.from({ length: comparison.maxCols }, (_, i) => getColumnLetter(i))];
      const dataRows = Array.from({ length: comparison.maxRows }).map((_, rowIdx) => [
        rowIdx + 1,
        ...Array.from({ length: comparison.maxCols }).map((_, colIdx) => String(sheet[rowIdx]?.[colIdx] ?? ''))
      ]);
      return [head, ...dataRows];
    };

    appendSheet(
      XLSX.utils.aoa_to_sheet(buildTable(comparison.sheet1)),
      `${selectedSheet1 || 'Sheet1'}`
    );
    appendSheet(
      XLSX.utils.aoa_to_sheet(buildTable(comparison.sheet2)),
      `${selectedSheet2 || 'Sheet2'}`
    );

    const sanitize = (name) => (name || 'file').replace(/[^a-z0-9]/gi, '_');
    const filename = `comparison_${sanitize(file1Name)}_${sanitize(file2Name)}.xlsx`;

    XLSX.writeFile(workbook, filename);
  };

  return (
    <div className="app-container">
      <div className="content-wrapper">
        <div className="card">
          <div className="header">
            <FileSpreadsheet className="header-icon" />
            <h1>Excel Comparison Tool</h1>
          </div>

          {error && (
            <div className="error-box">
              <AlertCircle className="error-icon" />
              <p>{error}</p>
            </div>
          )}

          <div className="upload-grid">
            <div className="upload-box">
              <label className="upload-label">
                <div className="upload-content">
                  <Upload className="upload-icon" />
                  <span className="upload-text">Upload First Excel File</span>
                  {file1Name && (
                    <span className="file-name">
                      <Check className="check-icon" /> {file1Name}
                    </span>
                  )}
                </div>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="file-input"
                  onChange={handleFile1Upload}
                />
              </label>
              {file1Data && (
                <select
                  className="sheet-select"
                  value={selectedSheet1}
                  onChange={(e) => setSelectedSheet1(e.target.value)}
                >
                  {file1Data.sheetNames.map(name => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              )}
            </div>

            <div className="upload-box">
              <label className="upload-label">
                <div className="upload-content">
                  <Upload className="upload-icon" />
                  <span className="upload-text">Upload Second Excel File</span>
                  {file2Name && (
                    <span className="file-name">
                      <Check className="check-icon" /> {file2Name}
                    </span>
                  )}
                </div>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="file-input"
                  onChange={handleFile2Upload}
                />
              </label>
              {file2Data && (
                <select
                  className="sheet-select"
                  value={selectedSheet2}
                  onChange={(e) => setSelectedSheet2(e.target.value)}
                >
                  {file2Data.sheetNames.map(name => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              )}
            </div>
          </div>

          {comparison && (
            <div className="results">
              <div className="summary-box">
                <div className="summary-header">
                  <h2>Comparison Summary</h2>
                  <button className="export-btn" onClick={handleDownloadResults}>
                    <Download className="btn-icon" />
                    Download Results
                  </button>
                </div>
                <div className="stats-grid">
                  <div className="stat-card">
                    <p className="stat-label">Total Cells</p>
                    <p className="stat-value">{comparison.totalCells}</p>
                  </div>
                  <div className="stat-card stat-success">
                    <p className="stat-label">Matching</p>
                    <p className="stat-value">{comparison.matchingCells}</p>
                  </div>
                  <div className="stat-card stat-error">
                    <p className="stat-label">Differences</p>
                    <p className="stat-value">{comparison.diffs.length}</p>
                  </div>
                </div>
              </div>

              <div className="comparison-section">
                <h3 className="section-title">
                  Side-by-Side Comparison
                  <span className="legend">(Green = Match, Red = Difference)</span>
                </h3>
                <div className="tables-grid">
                  <div>
                    <h4 className="table-header">
                      <FileSpreadsheet className="small-icon" />
                      {file1Name} - {selectedSheet1}
                    </h4>
                    <div className="table-container">
                      <table className="comparison-table">
                        <thead>
                          <tr>
                            <th className="row-header"></th>
                            {Array.from({ length: comparison.maxCols }).map((_, i) => (
                              <th key={i}>{getColumnLetter(i)}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {Array.from({ length: comparison.maxRows }).map((_, rowIdx) => (
                            <tr key={rowIdx}>
                              <td className="row-header">{rowIdx + 1}</td>
                              {Array.from({ length: comparison.maxCols }).map((_, colIdx) => {
                                const val = comparison.sheet1[rowIdx]?.[colIdx] ?? '';
                                const isDiff = comparison.diffs.some(d => d.row === rowIdx && d.col === colIdx);
                                return (
                                  <td key={colIdx} className={isDiff ? 'cell-diff' : 'cell-match'}>
                                    {String(val)}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <div>
                    <h4 className="table-header">
                      <FileSpreadsheet className="small-icon" />
                      {file2Name} - {selectedSheet2}
                    </h4>
                    <div className="table-container">
                      <table className="comparison-table">
                        <thead>
                          <tr>
                            <th className="row-header"></th>
                            {Array.from({ length: comparison.maxCols }).map((_, i) => (
                              <th key={i}>{getColumnLetter(i)}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {Array.from({ length: comparison.maxRows }).map((_, rowIdx) => (
                            <tr key={rowIdx}>
                              <td className="row-header">{rowIdx + 1}</td>
                              {Array.from({ length: comparison.maxCols }).map((_, colIdx) => {
                                const val = comparison.sheet2[rowIdx]?.[colIdx] ?? '';
                                const isDiff = comparison.diffs.some(d => d.row === rowIdx && d.col === colIdx);
                                return (
                                  <td key={colIdx} className={isDiff ? 'cell-diff' : 'cell-match'}>
                                    {String(val)}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>

              {comparison.diffs.length > 0 && (
                <div className="differences-section">
                  <h3 className="section-title">Differences List</h3>
                  <div className="diff-table-container">
                    <table className="diff-table">
                      <thead>
                        <tr>
                          <th>Cell</th>
                          <th>File 1</th>
                          <th>File 2</th>
                        </tr>
                      </thead>
                      <tbody>
                        {comparison.diffs.slice(0, 100).map((diff, idx) => (
                          <tr key={idx}>
                            <td className="cell-ref">{getColumnLetter(diff.col)}{diff.row + 1}</td>
                            <td>{String(diff.val1)}</td>
                            <td>{String(diff.val2)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {comparison.diffs.length > 100 && (
                      <p className="truncate-note">
                        Showing first 100 of {comparison.diffs.length} differences
                      </p>
                    )}
                  </div>
                </div>
              )}

              {comparison.diffs.length === 0 && (
                <div className="success-box">
                  <Check className="success-icon" />
                  <p>The selected sheets are identical!</p>
                </div>
              )}
            </div>
          )}

          {!file1Data && !file2Data && (
            <div className="empty-state">
              <FileSpreadsheet className="empty-icon" />
              <p>Upload two Excel files to begin comparison</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}