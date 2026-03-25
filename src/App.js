// Import necessary React hooks and libraries
import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx'; // Library for reading Excel files
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from 'docx'; // Library for creating Word documents
import { saveAs } from 'file-saver'; // Library for downloading files
import JSZip from 'jszip'; // Library for creating ZIP files
import './App.css';

function App() {
  // State management
  const [files, setFiles] = useState([]); // Array of uploaded Excel files with their data
  const [selectedFileIndex, setSelectedFileIndex] = useState(0); // Index of currently selected file for preview
  const [isDragging, setIsDragging] = useState(false); // Track drag-and-drop state for visual feedback
  const [transpose, setTranspose] = useState(true); // Whether to transpose rows/columns in output
  const [splitColumns, setSplitColumns] = useState(false); // Whether to split columns into separate documents
  const [showPreview, setShowPreview] = useState(true); // Whether to show the document preview
  const [selectedItem, setSelectedItem] = useState(''); // Index of selected item column for analysis
  const [itemAnalysis, setItemAnalysis] = useState(null); // Results of Y/N analysis for selected item

  // Reset item analysis whenever the active file changes
  useEffect(() => {
    setSelectedItem('');
    setItemAnalysis(null);
  }, [selectedFileIndex]);

  /**
   * Returns all column headers from the current file's first row
   */
  const getItemColumns = () => {
    if (!files || files.length === 0 || !files[selectedFileIndex]) return [];
    const data = files[selectedFileIndex].data;
    if (!data || data.length === 0) return [];
    return (data[0] || []).map((header, idx) => ({
      header: String(header !== undefined && header !== null ? header : `Column ${idx + 1}`),
      colIndex: idx,
    }));
  };

  /**
   * Scans the selected column for Y/N values across all data rows.
   * Flags a conflict if both Y and N are present.
   * @param {number} colIndex - Column index to analyse
   */
  const analyzeItem = (colIndex) => {
    const idx = parseInt(colIndex, 10);
    const data = files[selectedFileIndex].data;
    const yAddresses = [];
    const nAddresses = [];

    // Start from row 1 to skip the header row
    for (let row = 1; row < data.length; row++) {
      const address = data[row][0] !== undefined ? String(data[row][0]) : `Row ${row}`;
      const cellValue = data[row][idx] !== undefined
        ? String(data[row][idx]).trim().toLowerCase()
        : '';

      if (cellValue === 'y' || cellValue === 'yes') {
        yAddresses.push(address);
      } else if (cellValue === 'n' || cellValue === 'no') {
        nAddresses.push(address);
      }
      // All other values are intentionally ignored
    }

    setItemAnalysis({
      yAddresses,
      nAddresses,
      hasConflict: yAddresses.length > 0 && nAddresses.length > 0,
    });
  };

  /** Handler for the item dropdown */
  const handleItemSelect = (e) => {
    const val = e.target.value;
    setSelectedItem(val);
    if (val !== '') {
      analyzeItem(val);
    } else {
      setItemAnalysis(null);
    }
  };

  /**
   * Handles uploading and processing Excel files
   * Reads Excel files, converts them to JSON format, and adds them to state
   * @param {FileList} fileList - List of files to upload (from input or drag-drop)
   */
  const handleFilesUpload = async (fileList) => {
    if (!fileList || fileList.length === 0) return;

    const newFiles = [];
    
    // Process each file in the list
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      // Only process Excel files (.xlsx or .xls)
      if (file.name.match(/\.(xlsx|xls)$/i)) {
        try {
          // Read file as array buffer
          const data = await file.arrayBuffer();
          // Parse Excel workbook
          const workbook = XLSX.read(data, { type: 'array' });
          // Get first sheet from workbook
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          // Convert sheet to 2D array (rows and columns)
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          // Store file data with metadata
          newFiles.push({
            name: file.name.replace(/\.[^/.]+$/, ''), // Remove file extension
            data: jsonData,
            originalName: file.name
          });
        } catch (error) {
          console.error(`Error reading file ${file.name}:`, error);
        }
      }
    }
    
    // Add new files to state and select the first new file
    if (newFiles.length > 0) {
      setFiles(prev => [...prev, ...newFiles]);
      setSelectedFileIndex(files.length);
    }
  };

  /** Handler for file input change (Browse Files button) */
  const handleFileChange = (e) => {
    handleFilesUpload(e.target.files);
  };

  /** Handler for folder input change (Browse Folder button) */
  const handleFolderChange = (e) => {
    handleFilesUpload(e.target.files);
  };

  /** Handler when drag enters the drop zone */
  const handleDragEnter = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true); // Show visual feedback
  };

  /** Handler when drag leaves the drop zone */
  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false); // Remove visual feedback
  };

  /** Handler for drag over event (required for drop to work) */
  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  /** Handler when files are dropped into the drop zone */
  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const droppedFiles = e.dataTransfer.files;
    if (droppedFiles.length > 0) {
      handleFilesUpload(droppedFiles);
    }
  };

  /**
   * Transposes a 2D array (swaps rows and columns)
   * Example: [[1,2,3], [4,5,6]] becomes [[1,4], [2,5], [3,6]]
   * @param {Array<Array>} data - 2D array to transpose
   * @returns {Array<Array>} Transposed 2D array
   */
  const transposeData = (data) => {
    if (!data || data.length === 0) return data;
    
    // Find the maximum number of columns across all rows
    const maxCols = Math.max(...data.map(row => row.length));
    const transposed = [];
    
    // Iterate through each column index
    for (let col = 0; col < maxCols; col++) {
      const newRow = [];
      // Iterate through each row and extract the value at current column
      for (let row = 0; row < data.length; row++) {
        newRow.push(data[row][col] !== undefined ? data[row][col] : '');
      }
      transposed.push(newRow);
    }
    
    return transposed;
  };

  /**
   * Splits data into multiple column pairs (first column + each subsequent column)
   * Example: [[A,B,C], [1,2,3]] becomes [[[A,B], [1,2]], [[A,C], [1,3]]]
   * @param {Array<Array>} data - 2D array to split
   * @returns {Array<Array<Array>>} Array of column pairs
   */
  const splitIntoColumnPairs = (data) => {
    if (!data || data.length === 0) return [data];
    
    const maxCols = Math.max(...data.map(row => row.length));
    if (maxCols <= 1) return [data]; // Not enough columns to split
    
    const pairs = [];
    
    // Create a pair for each column after the first
    for (let col = 1; col < maxCols; col++) {
      const pair = data.map(row => [
        row[0] !== undefined ? row[0] : '', // First column (labels)
        row[col] !== undefined ? row[col] : '' // Current value column
      ]);
      pairs.push(pair);
    }
    
    return pairs;
  };

  /**
   * Creates a formatted packing sheet Word document from item analysis results.
   * Format: "ITEM: These addresses DO WANT" or "ITEM: These addresses DO NOT WANT"
   * followed by one address per line.
   */
  const createItemAnalysisDocument = async (itemName, isYDoc, addresses) => {
    const label = isYDoc ? 'These addresses DO WANT' : 'These addresses DO NOT WANT';
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            children: [new TextRun({ text: `${itemName.toUpperCase()}: ${label}`, bold: true, size: 32 })],
            spacing: { after: 200 },
          }),
          ...addresses.map(addr => new Paragraph({
            children: [new TextRun({ text: addr, size: 24 })],
            spacing: { after: 100 },
          })),
        ],
      }],
    });
    return await Packer.toBlob(doc);
  };

  /** Downloads the item analysis result as a formatted Word document */
  const downloadItemAnalysisDocument = async () => {
    const itemHeader = getItemColumns().find(c => c.colIndex === parseInt(selectedItem, 10));
    const itemName = itemHeader ? itemHeader.header : 'Item';
    const isYDoc = itemAnalysis.yAddresses.length > 0;
    const addresses = isYDoc ? itemAnalysis.yAddresses : itemAnalysis.nAddresses;
    try {
      const blob = await createItemAnalysisDocument(itemName, isYDoc, addresses);
      const suffix = isYDoc ? 'DO_WANT' : 'DO_NOT_WANT';
      saveAs(blob, `${itemName}_${suffix}.docx`);
    } catch (error) {
      console.error('Error creating document:', error);
      alert('Error creating document. Please try again.');
    }
  };

  /**
   * Creates a Word document (.docx) from Excel data
   * @param {Object} fileData - File object containing Excel data
   * @param {boolean} transpose - Whether to transpose the data
   * @param {string} suffix - Optional suffix to add to the title (e.g., family name)
   * @returns {Blob} Word document as a blob
   */
  const createWordDocument = async (fileData, transpose, suffix = '') => {
    // Apply transpose if enabled
    const dataToExport = transpose ? transposeData(fileData.data) : fileData.data;

    // Convert each row of data into Word table rows
    const tableRows = dataToExport.map((row) => {
      return new TableRow({
        children: row.map((cell) => {
          // Each cell contains a paragraph with text
          return new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: cell !== null && cell !== undefined ? String(cell) : '',
                    size: 20, // Font size in half-points (20 = 10pt)
                  }),
                ],
              }),
            ],
          });
        }),
      });
    });

    // Create Word document with title and table
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            // Title paragraph
            new Paragraph({
              children: [
                new TextRun({
                  text: `Packing Sheet${suffix ? ' ' + suffix : ''}`,
                  bold: true,
                  size: 32, // 16pt font
                }),
              ],
              spacing: {
                after: 200, // Space after title
              },
            }),
            // Data table
            new Table({
              rows: tableRows,
              width: {
                size: 100,
                type: 'pct', // 100% width
              },
            }),
          ],
        },
      ],
    });

    // Convert document to blob for download
    return await Packer.toBlob(doc);
  };

  /**
   * Downloads the currently selected file as a Word document
   * If splitColumns is enabled, downloads multiple documents (one per column pair)
   */
  const downloadWordFile = async () => {
    if (!files || files.length === 0) {
      alert('Please upload an Excel file first!');
      return;
    }

    const currentFile = files[selectedFileIndex];
    if (!currentFile) {
      alert('No file selected!');
      return;
    }

    try {
      const dataToExport = transpose ? transposeData(currentFile.data) : currentFile.data;
      
      if (splitColumns) {
        // Split into column pairs and download each
        const columnPairs = splitIntoColumnPairs(dataToExport);
        
        for (let i = 0; i < columnPairs.length; i++) {
          const familyName = columnPairs[i][0]?.[1] || `column${i + 2}`;
          const blob = await createWordDocument({ ...currentFile, data: columnPairs[i] }, false, familyName);
          saveAs(blob, `${currentFile.name}_${familyName}.docx`);
          
          // Add delay between downloads
          if (i < columnPairs.length - 1) {
            await new Promise(resolve => setTimeout(resolve, 300));
          }
        }
        
        alert(`Successfully downloaded ${columnPairs.length} Word documents!`);
      } else {
        // Download single document
        const blob = await createWordDocument(currentFile, transpose);
        saveAs(blob, `${currentFile.name || 'converted'}.docx`);
      }
    } catch (error) {
      console.error('Error creating Word document:', error);
      alert('Error creating Word document. Please try again.');
    }
  };

  /**
   * Downloads all loaded files as separate Word documents
   * Adds delay between downloads to prevent browser blocking
   * If splitColumns is enabled, creates multiple documents per file
   */
  const downloadAllWordFiles = async () => {
    if (!files || files.length === 0) {
      alert('Please upload Excel files first!');
      return;
    }

    try {
      let totalDocs = 0;
      
      // Download each file with a small delay to avoid browser blocking
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const dataToExport = transpose ? transposeData(file.data) : file.data;
        
        if (splitColumns) {
          // Split into column pairs and download each
          const columnPairs = splitIntoColumnPairs(dataToExport);
          
          for (let j = 0; j < columnPairs.length; j++) {
            const familyName = columnPairs[j][0]?.[1] || `column${j + 2}`;
            const blob = await createWordDocument({ ...file, data: columnPairs[j] }, false, familyName);
            saveAs(blob, `${file.name}_${familyName}.docx`);
            totalDocs++;
            
            // Add a 300ms delay between downloads
            await new Promise(resolve => setTimeout(resolve, 300));
          }
        } else {
          // Download single document per file
          const blob = await createWordDocument(file, transpose);
          saveAs(blob, `${file.name || `converted_${i + 1}`}.docx`);
          totalDocs++;
          
          // Add a 300ms delay between downloads to ensure they all trigger
          if (i < files.length - 1) {
            await new Promise(resolve => setTimeout(resolve, 300));
          }
        }
      }
      
      alert(`Successfully downloaded ${totalDocs} Word document${totalDocs > 1 ? 's' : ''}!`);
    } catch (error) {
      console.error('Error creating Word documents:', error);
      alert('Error creating Word documents. Please try again.');
    }
  };

  /**
   * Downloads all files as Word documents packaged in a single ZIP file
   * Useful for downloading many files at once without browser blocking
   * If splitColumns is enabled, includes all column-split documents in the ZIP
   */
  const downloadZipFile = async () => {
    if (!files || files.length === 0) {
      alert('Please upload Excel files first!');
      return;
    }

    try {
      const zip = new JSZip();
      
      // Create a folder in the zip for organization
      const folder = zip.folder('converted_files');
      
      let totalDocs = 0;
      
      // Convert each Excel file to Word and add to ZIP
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const dataToExport = transpose ? transposeData(file.data) : file.data;
        
        if (splitColumns) {
          // Split into column pairs and add each to ZIP
          const columnPairs = splitIntoColumnPairs(dataToExport);
          
          for (let j = 0; j < columnPairs.length; j++) {
            const familyName = columnPairs[j][0]?.[1] || `column${j + 2}`;
            const blob = await createWordDocument({ ...file, data: columnPairs[j] }, false, familyName);
            folder.file(`${file.name}_${familyName}.docx`, blob);
            totalDocs++;
          }
        } else {
          // Add single document to ZIP
          const blob = await createWordDocument(file, transpose);
          folder.file(`${file.name || `converted_${i + 1}`}.docx`, blob);
          totalDocs++;
        }
      }
      
      // Generate the zip file as a blob
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      saveAs(zipBlob, 'converted_files.zip');
      
      alert(`Successfully created zip file with ${totalDocs} Word document${totalDocs > 1 ? 's' : ''}!`);
    } catch (error) {
      console.error('Error creating zip file:', error);
      alert('Error creating zip file. Please try again.');
    }
  };

  return (
    <div className="App">
      <div className="container">
        {/* Header */}
        <h1>Packing Sheet Generator</h1>
        <p className="subtitle">Upload your Excel file and download it as a Word document</p>

        {/* Main content area - side by side layout */}
        <div className={`main-content ${files.length === 0 || !showPreview ? 'centered' : ''}`}>
          {/* Left side - Upload section */}
          <div className="left-section">
            {/* Upload Area - supports drag & drop, file selection, and folder selection */}
            <div
              className={`upload-area ${isDragging ? 'dragging' : ''}`}
              onDragEnter={handleDragEnter}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <div className="upload-icon">📊</div>
              <h3>Drag & Drop Excel Files or Folder Here</h3>
              <p>or</p>
              <div className="upload-buttons">
                {/* Browse Files button - allows selecting multiple individual files */}
                <label htmlFor="file-upload" className="file-label">
                  Browse Files
                </label>
                {/* Browse Folder button - selects all files in a folder */}
                <label htmlFor="folder-upload" className="file-label folder-label">
                  Browse Folder
                </label>
              </div>
              {/* Hidden file input for individual file selection */}
              <input
                id="file-upload"
                type="file"
                accept=".xlsx,.xls"
                multiple
                onChange={handleFileChange}
                style={{ display: 'none' }}
              />
              {/* Hidden file input for folder selection */}
              <input
                id="folder-upload"
                type="file"
                webkitdirectory=""
                directory=""
                onChange={handleFolderChange}
                style={{ display: 'none' }}
              />
              {/* File count indicator */}
              {files.length > 0 && (
                <div className="file-info">
                  <p>✓ {files.length} file{files.length > 1 ? 's' : ''} loaded</p>
                </div>
              )}
            </div>

            {/* File tabs - show when files are loaded */}
            {files.length > 0 && (
              <div className="files-list">
                <h3>Loaded Files ({files.length})</h3>
                <div className="file-tabs">
                  {files.map((file, index) => (
                    <button
                      key={index}
                      className={`file-tab ${index === selectedFileIndex ? 'active' : ''}`}
                      onClick={() => setSelectedFileIndex(index)}
                    >
                      {file.originalName}
                      {/* Remove file button (X) */}
                      <span
                        className="remove-file"
                        onClick={(e) => {
                          e.stopPropagation(); // Prevent tab selection when clicking X
                          const newFiles = files.filter((_, i) => i !== index);
                          setFiles(newFiles);
                          // Adjust selected index if needed
                          if (selectedFileIndex >= newFiles.length) {
                            setSelectedFileIndex(Math.max(0, newFiles.length - 1));
                          }
                        }}
                      >
                        ×
                      </span>
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* Options - show when files are loaded */}
            {files.length > 0 && files[selectedFileIndex] && (
              <div className="options-section">
                <h3>Options</h3>
                {/* Transpose toggle checkbox */}
                <div className="transpose-toggle">
                  <label>
                    <input
                      type="checkbox"
                      checked={transpose}
                      onChange={(e) => setTranspose(e.target.checked)}
                    />
                    <span>Transpose rows and columns</span>
                  </label>
                </div>
                {/* Split columns toggle checkbox */}
                <div className="transpose-toggle">
                  <label>
                    <input
                      type="checkbox"
                      checked={splitColumns}
                      onChange={(e) => setSplitColumns(e.target.checked)}
                    />
                    <span>Split into separate documents (first column + each other column)</span>
                  </label>
                </div>
                {/* Show preview toggle checkbox */}
                <div className="transpose-toggle">
                  <label>
                    <input
                      type="checkbox"
                      checked={showPreview}
                      onChange={(e) => setShowPreview(e.target.checked)}
                    />
                    <span>Show document preview</span>
                  </label>
                </div>
              </div>
            )}
          </div>

          {/* Right side - Preview section */}
          {files.length > 0 && files[selectedFileIndex] && showPreview && (
            <div className="right-section">
              <div className="preview-section">
                <h3>Document Preview - {files[selectedFileIndex].originalName}</h3>
                {/* Document preview styled like Word */}
                <div className="document-preview">
                  {itemAnalysis && !itemAnalysis.hasConflict ? (
                    // Formatted item analysis preview
                    (() => {
                      const itemHeader = getItemColumns().find(c => c.colIndex === parseInt(selectedItem, 10));
                      const itemName = itemHeader ? itemHeader.header.toUpperCase() : '';
                      const isYDoc = itemAnalysis.yAddresses.length > 0;
                      const addresses = isYDoc ? itemAnalysis.yAddresses : itemAnalysis.nAddresses;
                      const label = isYDoc ? 'These addresses DO WANT' : 'These addresses DO NOT WANT';
                      return (
                        <div className="document-page">
                          <h2 className="document-title">{itemName}: {label}</h2>
                          {addresses.map((addr, i) => (
                            <p key={i} className="address-preview-item">{addr}</p>
                          ))}
                        </div>
                      );
                    })()
                  ) : (
                    // Raw spreadsheet table preview
                    (() => {
                      const dataToExport = transpose ? transposeData(files[selectedFileIndex].data) : files[selectedFileIndex].data;
                      const previews = splitColumns ? splitIntoColumnPairs(dataToExport) : [dataToExport];
                      return previews.map((previewData, idx) => {
                        const familyName = splitColumns && previewData[0]?.[1] ? previewData[0][1] : '';
                        return (
                          <div key={idx} className="document-page">
                            <h2 className="document-title">Packing Sheet{familyName ? ` ${familyName}` : ''}</h2>
                            <div className="table-container">
                              <table className="document-table">
                                <tbody>
                                  {previewData.map((row, rowIndex) => (
                                    <tr key={rowIndex}>
                                      {row.map((cell, cellIndex) => (
                                        <td key={cellIndex}>
                                          {cell !== null && cell !== undefined ? String(cell) : ''}
                                        </td>
                                      ))}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        );
                      });
                    })()
                  )}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Item Analysis Section */}
        {files.length > 0 && files[selectedFileIndex] && (
          <div className="item-analysis-section">
            <h3>Item Analysis</h3>
            <div className="item-select-row">
              <label htmlFor="item-select">Select an item:</label>
              <select id="item-select" value={selectedItem} onChange={handleItemSelect}>
                <option value="">-- Select an item --</option>
                {getItemColumns().map(({ header, colIndex }) => (
                  <option key={colIndex} value={colIndex}>{header}</option>
                ))}
              </select>
            </div>

            {itemAnalysis && (
              <div className="analysis-results-container">
                {itemAnalysis.hasConflict && (
                  <div className="conflict-error">
                    ⚠️ Error: This item has both &quot;Yes&quot; and &quot;No&quot; values in the same column.
                  </div>
                )}
                {!itemAnalysis.hasConflict && (
                  <button className="download-btn analysis-download-btn" onClick={downloadItemAnalysisDocument}>
                    📄 Download Packing Sheet
                  </button>
                )}
                <div className="analysis-results">
                  <div className="address-list y-addresses">
                    <h4>✅ DO Want ({itemAnalysis.yAddresses.length})</h4>
                    {itemAnalysis.yAddresses.length > 0 ? (
                      <ul>
                        {itemAnalysis.yAddresses.map((addr, i) => (
                          <li key={i}>{addr}</li>
                        ))}
                      </ul>
                    ) : (
                      <p className="no-entries">No entries</p>
                    )}
                  </div>
                  <div className="address-list n-addresses">
                    <h4>❌ DO NOT Want ({itemAnalysis.nAddresses.length})</h4>
                    {itemAnalysis.nAddresses.length > 0 ? (
                      <ul>
                        {itemAnalysis.nAddresses.map((addr, i) => (
                          <li key={i}>{addr}</li>
                        ))}
                      </ul>
                    ) : (
                      <p className="no-entries">No entries</p>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Download Buttons */}
        <div className="download-buttons-container">
          {/* Download current file button */}
          <button
            className="download-btn"
            onClick={downloadWordFile}
            disabled={files.length === 0}
          >
            📄 Download Current File as Word Document
          </button>
          
          {/* Additional buttons when multiple files are loaded */}
          {files.length > 1 && (
            <>
              {/* Download all files separately */}
              <button
                className="download-btn download-all-btn"
                onClick={downloadAllWordFiles}
                disabled={files.length === 0}
              >
                📦 Download All Files Separately ({files.length})
              </button>
              
              {/* Download all files as ZIP */}
              <button
                className="download-btn download-zip-btn"
                onClick={downloadZipFile}
                disabled={files.length === 0}
              >
                🗜️ Download All as ZIP ({files.length} files)
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

export default App;
