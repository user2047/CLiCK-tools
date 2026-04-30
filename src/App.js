// Import necessary React hooks and libraries
import React, { useState } from 'react';
import * as XLSX from 'xlsx'; // Library for reading Excel files
import { AlignmentType, Document, Packer, Paragraph, Table, TableBorders, TableCell, TableRow, TextRun } from 'docx'; // Library for creating Word documents
import { saveAs } from 'file-saver'; // Library for downloading files
import JSZip from 'jszip'; // Library for creating ZIP files
import './App.css';

const LABEL_ALIASES = [
  ['family', 'family name', 'household', 'household name', 'client', 'client name', 'name'],
  ['num people', 'number of people', 'people', 'household size', 'adults', 'children', 'kids', 'seniors'],
  ['misc wants', 'wants', 'request', 'requests', 'requested items', 'notes'],
  ["misc don't include", 'misc dont include', 'do not include', "don't include", 'dont include', 'restrictions', 'allergies'],
  ['fruits', 'fruit', 'vegetables', 'veggies', 'fruits veggies', 'fruits and veggies', 'fruits & veggies'],
  ['dried goods', 'canned goods', 'dried canned goods', 'dried/canned goods'],
  ['snacks', 'snack'],
  ['cooking items', 'cooking', 'pantry staples'],
  ['meat', 'protein', 'dairy', 'bread', 'eggs', 'hygiene', 'diapers']
];

const MESA_ROUTE_ALIASES = [
  [
    'first',
    'first name',
    'nombre',
    'last',
    'last name',
    'apellido',
    'recipient name',
    'full name',
    'delivery address',
    'address',
    'where to leave food bag',
    'language',
    'preferred method of contact',
    'phone',
    'phone number',
    'email'
  ],
  [
    'misc preferences',
    'dietary restrictions',
    'pet food',
    'baby food',
    'clothes',
    'coffee'
  ]
];

const HEADER_ALIASES = [...LABEL_ALIASES, ...MESA_ROUTE_ALIASES];
const HEADER_WORDS = new Set(HEADER_ALIASES.flatMap(group => group.flatMap(label => label.split(' '))));

const normalizeCell = (value) => {
  if (value === null || value === undefined) return '';
  return String(value).replace(/\s+/g, ' ').trim();
};

const normalizeKey = (value) => {
  return normalizeCell(value)
    .toLowerCase()
    .replace(/[\u2019]/g, "'")
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9']+/g, ' ')
    .trim();
};

const scoreHeaderCell = (value) => {
  const key = normalizeKey(value).replace(/'/g, '');
  if (!key) return 0;

  for (const group of HEADER_ALIASES) {
    if (group.map(alias => alias.replace(/'/g, '')).includes(key)) {
      return 6;
    }
  }

  return key
    .split(' ')
    .filter(word => HEADER_WORDS.has(word))
    .length;
};

const scoreHeaderLine = (cells) => {
  return cells.reduce((score, cell) => score + scoreHeaderCell(cell), 0);
};

const trimMatrix = (rows) => {
  const normalizedRows = rows.map(row => (Array.isArray(row) ? row : []).map(normalizeCell));
  const nonEmptyRows = normalizedRows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => row.some(cell => cell !== ''));

  if (nonEmptyRows.length === 0) return [];

  const firstRow = nonEmptyRows[0].index;
  const lastRow = nonEmptyRows[nonEmptyRows.length - 1].index;
  const maxCols = normalizedRows.reduce((max, row) => Math.max(max, row.length), 0);
  const usedCols = [];

  for (let col = 0; col < maxCols; col++) {
    const hasValue = normalizedRows
      .slice(firstRow, lastRow + 1)
      .some(row => normalizeCell(row[col]) !== '');
    if (hasValue) usedCols.push(col);
  }

  if (usedCols.length === 0) return [];

  const firstCol = usedCols[0];
  const lastCol = usedCols[usedCols.length - 1];

  return normalizedRows
    .slice(firstRow, lastRow + 1)
    .map(row => {
      const trimmedRow = [];
      for (let col = firstCol; col <= lastCol; col++) {
        trimmedRow.push(normalizeCell(row[col]));
      }
      return trimmedRow;
    })
    .filter(row => row.some(cell => cell !== ''));
};

const findTableStartIndex = (rows) => {
  let bestIndex = 0;
  let bestScore = -1;
  const rowsToScan = Math.min(rows.length, 30);

  for (let index = 0; index < rowsToScan; index++) {
    const row = rows[index];
    const filledCells = row.filter(cell => cell !== '').length;
    if (filledCells < 2) continue;

    const score = (scoreHeaderLine(row) * 3) + Math.min(filledCells, 12);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  }

  if (bestScore >= 8) return bestIndex;
  return rows.findIndex(row => row.filter(cell => cell !== '').length >= 2);
};

const transposeMatrix = (data) => {
  if (!data || data.length === 0) return data;

  const maxCols = Math.max(...data.map(row => row.length));
  const transposed = [];

  for (let col = 0; col < maxCols; col++) {
    const newRow = [];
    for (let row = 0; row < data.length; row++) {
      newRow.push(data[row][col] !== undefined ? data[row][col] : '');
    }
    transposed.push(newRow);
  }

  return transposed;
};

const prepareWorksheetRows = (rows) => {
  const trimmedRows = trimMatrix(rows);
  if (trimmedRows.length === 0) return [];

  const tableStart = findTableStartIndex(trimmedRows);
  const tableRows = tableStart >= 0 ? trimmedRows.slice(tableStart) : trimmedRows;
  return trimMatrix(tableRows);
};

const detectTransposePreference = (data) => {
  if (!data || data.length === 0) return true;

  const firstRowScore = scoreHeaderLine(data[0] || []);
  const firstColumnScore = scoreHeaderLine(data.map(row => row[0] || ''));

  return firstRowScore >= firstColumnScore;
};

const shouldSplitByDefault = (data, shouldTranspose) => {
  const documentData = shouldTranspose ? transposeMatrix(data) : data;
  return documentData.length > 0 && (documentData[0] || []).filter(cell => cell !== '').length > 2;
};

const getWorksheetScore = (data) => {
  if (!data || data.length === 0) return 0;

  const filledCells = data.reduce(
    (count, row) => count + row.filter(cell => cell !== '').length,
    0
  );

  return (scoreHeaderLine(data[0] || []) * 5)
    + (scoreHeaderLine(data.map(row => row[0] || '')) * 2)
    + Math.min(filledCells, 200);
};

const extractBestWorksheet = (workbook) => {
  const candidates = workbook.SheetNames
    .map(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        blankrows: false,
        raw: false,
        dateNF: 'm/d/yyyy'
      });
      const data = prepareWorksheetRows(rows);
      const transposeByDefault = detectTransposePreference(data);
      const packingConfig = getPackingSheetConfig(data);

      return {
        sheetName,
        data,
        transposeByDefault,
        splitByDefault: shouldSplitByDefault(data, transposeByDefault),
        packingConfig,
        packingScore: getPackingSheetScore(data),
        score: getWorksheetScore(data)
      };
    })
    .filter(candidate => candidate.data.length > 0);

  candidates.sort((a, b) => (b.packingScore - a.packingScore) || (b.score - a.score));
  return candidates[0] || null;
};

const sanitizeFilePart = (value, fallback = 'converted') => {
  const cleaned = normalizeCell(value)
    .replace(/[<>:"/\\|?*]+/g, '-')
    .replace(/\.+$/g, '')
    .trim();

  return (cleaned || fallback).slice(0, 80);
};

const escapeHtml = (value) => {
  return normalizeCell(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
};

const LIST_MODES = {
  white: {
    label: 'White list',
    title: 'Getting'
  },
  black: {
    label: 'Black list',
    title: 'Not getting'
  }
};

const ADDRESS_HEADERS = [
  'address',
  'delivery address',
  'delivery address direccion',
  'address direccion',
  'direccion'
];

const HOUSEHOLD_SIZE_HEADERS = [
  'household size',
  'family size',
  'num people',
  'number of people',
  'people'
];

const NON_ITEM_HEADER_PARTS = [
  'start date',
  'end date',
  'first',
  'nombre',
  'last',
  'apellido',
  'name',
  'address',
  'leave food bag',
  'where to leave',
  'language',
  'contact',
  'mailchimp',
  'phone',
  'email',
  'preferences',
  'dietary restrictions',
  'misc notes',
  'notes',
  'household',
  'adults',
  'children',
  'latinx',
  'elderly',
  'disabled',
  'gender',
  'ethnicity',
  'birth',
  'waitlist',
  'snap',
  'hear',
  'apply',
  'reason',
  'survey',
  'status',
  'allergies',
  'enough food',
  'toss out food',
  'program comments',
  'other problems'
];

const NEGATIVE_VALUES = new Set([
  'n',
  'no',
  'none',
  'false',
  '0',
  'na',
  'n a',
  'n/a',
  'not applicable'
]);

const hasAnyHeaderMatch = (key, aliases) => {
  return aliases.some(alias => key === alias || key.includes(alias));
};

const findHeaderIndex = (headers, aliases) => {
  return headers.findIndex(header => hasAnyHeaderMatch(normalizeKey(header).replace(/'/g, ''), aliases));
};

const isItemHeader = (header) => {
  const key = normalizeKey(header).replace(/'/g, '');
  if (!key) return false;
  return !NON_ITEM_HEADER_PARTS.some(part => key === part || key.includes(part));
};

const isGettingItem = (value) => {
  const raw = normalizeCell(value);
  if (!raw) return false;

  const key = normalizeKey(raw).replace(/'/g, '');
  if (!key || NEGATIVE_VALUES.has(key)) return false;
  if (key.startsWith('no ') || key.startsWith('do not ') || key.startsWith('dont ')) return false;

  return true;
};

const getPackingSheetConfig = (data) => {
  if (!data || data.length < 2) {
    return {
      addressIndex: -1,
      householdSizeIndex: -1,
      itemColumns: []
    };
  }

  const headers = data[0] || [];
  const addressIndex = findHeaderIndex(headers, ADDRESS_HEADERS);
  const householdSizeIndex = findHeaderIndex(headers, HOUSEHOLD_SIZE_HEADERS);
  const firstItemIndex = addressIndex >= 0 ? addressIndex + 1 : 0;

  const itemColumns = headers
    .map((header, index) => ({ index, label: normalizeCell(header) }))
    .filter(({ index, label }) => {
      if (index < firstItemIndex || index === addressIndex || index === householdSizeIndex) return false;
      return isItemHeader(label);
    });

  return {
    addressIndex,
    householdSizeIndex,
    itemColumns
  };
};

const getPackingSheetScore = (data) => {
  const config = getPackingSheetConfig(data);
  if (config.addressIndex < 0 || config.itemColumns.length === 0) return 0;

  const rowCount = Math.max(data.length - 1, 0);
  return 1000 + (config.itemColumns.length * 25) + Math.min(rowCount, 200);
};

const getPackingSheets = (fileData, listMode) => {
  const config = fileData?.packingConfig || getPackingSheetConfig(fileData?.data || []);
  if (!fileData || config.addressIndex < 0 || config.itemColumns.length === 0) return [];

  const rows = fileData.data.slice(1);

  return config.itemColumns.map(itemColumn => {
    const recipients = rows
      .map(row => {
        const address = normalizeCell(row[config.addressIndex]);
        const householdSize = config.householdSizeIndex >= 0
          ? normalizeCell(row[config.householdSizeIndex])
          : '';
        const itemValue = normalizeCell(row[itemColumn.index]);
        const gettingItem = isGettingItem(itemValue);

        return {
          address,
          householdSize,
          itemValue,
          gettingItem
        };
      })
      .filter(recipient => recipient.address !== '')
      .filter(recipient => listMode === 'white' ? recipient.gettingItem : !recipient.gettingItem);

    const tableRows = [
      ['Address', 'Household Size', 'Details'],
      ...recipients.map(recipient => [
        recipient.address,
        recipient.householdSize,
        recipient.itemValue || (listMode === 'black' ? 'Not marked' : '')
      ])
    ];

    return {
      itemKey: `${itemColumn.index}-${itemColumn.label}`,
      itemName: itemColumn.label,
      title: `${itemColumn.label} - ${LIST_MODES[listMode].label}`,
      recipients,
      rows: tableRows
    };
  });
};

function App() {
  // State management
  const [files, setFiles] = useState([]); // Array of uploaded Excel files with their data
  const [selectedFileIndex, setSelectedFileIndex] = useState(0); // Index of currently selected file for preview
  const [isDragging, setIsDragging] = useState(false); // Track drag-and-drop state for visual feedback
  const [listMode, setListMode] = useState('white'); // Whether packing sheets list recipients getting or not getting each item
  const [showPreview, setShowPreview] = useState(true); // Whether to show the document preview
  const [includeBorders, setIncludeBorders] = useState(true); // Whether output tables include borders
  const [selectedPackingSheetKeys, setSelectedPackingSheetKeys] = useState({}); // Per item/page output selection

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
          // Parse Excel workbook and extract the best data table. Newer MESA
          // exports can include title/metadata rows or a summary sheet before
          // the actual packing data.
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const extracted = extractBestWorksheet(workbook);

          if (!extracted) continue;
          
          // Store file data with metadata
          newFiles.push({
            id: `${Date.now()}-${i}-${file.name}`,
            name: sanitizeFilePart(file.name.replace(/\.[^/.]+$/, '')), // Remove file extension
            data: extracted.data,
            originalName: workbook.SheetNames.length > 1
              ? `${file.name} - ${extracted.sheetName}`
              : file.name,
            sheetName: extracted.sheetName,
            transposeByDefault: extracted.transposeByDefault,
            splitByDefault: extracted.splitByDefault,
            packingConfig: extracted.packingConfig
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
   * Creates a Word document (.docx) from packing sheet rows.
   * @param {Object} packingSheet - Packing sheet to include
   * @returns {Blob} Word document as a blob
   */
  const createTableDocument = async (packingSheet) => {
    const { itemName, rows } = packingSheet;
    // Convert each row of data into Word table rows
    const tableRows = rows.map((row) => {
      return new TableRow({
        children: row.map((cell) => {
          // Each cell contains a paragraph with text
          return new TableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: cell !== null && cell !== undefined ? String(cell) : '',
                    bold: row === rows[0],
                    size: 20, // Font size in half-points (20 = 10pt)
                  }),
                ],
              }),
            ],
            borders: includeBorders ? undefined : TableBorders.NONE,
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
                  text: itemName,
                  bold: true,
                  size: 32, // 16pt font
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                after: 80,
              },
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: LIST_MODES[listMode].label,
                  size: 20,
                }),
              ],
              alignment: AlignmentType.CENTER,
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
              borders: includeBorders ? undefined : TableBorders.NONE,
            }),
          ],
        },
      ],
    });

    // Convert document to blob for download
    return await Packer.toBlob(doc);
  };

  const currentFile = files[selectedFileIndex];
  const packingSheets = getPackingSheets(currentFile, listMode);
  const getSelectionKey = (packingSheet) => {
    const fileKey = currentFile?.id || currentFile?.originalName || 'current-file';
    return `${fileKey}:${listMode}:${packingSheet.itemKey || packingSheet.itemName}`;
  };
  const isPackingSheetSelected = (packingSheet) => selectedPackingSheetKeys[getSelectionKey(packingSheet)] !== false;
  const selectedPackingSheets = packingSheets.filter(isPackingSheetSelected);
  const selectedPackingSheetCount = selectedPackingSheets.length;

  const setAllPackingSheetsSelected = (selected) => {
    setSelectedPackingSheetKeys(prev => {
      const next = { ...prev };
      packingSheets.forEach(packingSheet => {
        next[getSelectionKey(packingSheet)] = selected;
      });
      return next;
    });
  };

  const setPackingSheetSelected = (packingSheet, selected) => {
    setSelectedPackingSheetKeys(prev => ({
      ...prev,
      [getSelectionKey(packingSheet)]: selected
    }));
  };

  const downloadAllPackingSheets = async () => {
    if (!currentFile || packingSheets.length === 0) {
      alert('Please upload a MESA Excel file with address and item columns first!');
      return;
    }

    if (selectedPackingSheets.length === 0) {
      alert('Please select at least one item page to download.');
      return;
    }

    try {
      const zip = new JSZip();
      const folder = zip.folder(`${sanitizeFilePart(currentFile.name)}_${listMode}_packing_sheets`);

      for (const packingSheet of selectedPackingSheets) {
        const blob = await createTableDocument(packingSheet);
        folder.file(`${sanitizeFilePart(packingSheet.itemName)}_${listMode}.docx`, blob);
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      saveAs(zipBlob, `${sanitizeFilePart(currentFile.name)}_${listMode}_packing_sheets.zip`);
      alert(`Successfully created ${selectedPackingSheets.length} packing sheet${selectedPackingSheets.length > 1 ? 's' : ''}!`);
    } catch (error) {
      console.error('Error creating packing sheets:', error);
      alert('Error creating packing sheets. Please try again.');
    }
  };

  const printSelectedPackingSheets = () => {
    if (!currentFile || packingSheets.length === 0) {
      alert('Please upload a MESA Excel file with address and item columns first!');
      return;
    }

    if (selectedPackingSheets.length === 0) {
      alert('Please select at least one item page to print.');
      return;
    }

    const borderCss = includeBorders
      ? 'td { border: 1px solid #333; }'
      : 'td { border: none; }';
    const pagesHtml = selectedPackingSheets.map(packingSheet => `
      <section class="page">
        <h1>${escapeHtml(packingSheet.itemName)}</h1>
        <p class="list-mode">${escapeHtml(LIST_MODES[listMode].label)}</p>
        <table>
          <tbody>
            ${packingSheet.rows.map((row, rowIndex) => `
              <tr class="${rowIndex === 0 ? 'header-row' : ''}">
                ${row.map(cell => `<td>${escapeHtml(cell)}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </section>
    `).join('');

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      alert('Please allow pop-ups to print packing sheets.');
      return;
    }

    printWindow.document.open();
    printWindow.document.write(`
      <!doctype html>
      <html>
        <head>
          <title>${escapeHtml(currentFile.name)} Packing Sheets</title>
          <style>
            * { box-sizing: border-box; }
            body {
              margin: 0;
              color: #000;
              font-family: Arial, Helvetica, sans-serif;
              background: white;
            }
            .page {
              min-height: 100vh;
              padding: 0.5in;
              page-break-after: always;
            }
            .page:last-child {
              page-break-after: auto;
            }
            h1 {
              margin: 0 0 6px 0;
              text-align: center;
              font-size: 18pt;
            }
            .list-mode {
              margin: 0 0 18px 0;
              text-align: center;
              font-size: 10pt;
            }
            table {
              width: 100%;
              border-collapse: collapse;
            }
            td {
              padding: 6px 8px;
              font-size: 10pt;
              vertical-align: top;
            }
            ${borderCss}
            .header-row td {
              font-weight: 700;
              background: #eef2f7;
            }
            @page {
              margin: 0.5in;
            }
          </style>
        </head>
        <body>${pagesHtml}</body>
      </html>
    `);
    printWindow.document.close();
    printWindow.focus();
    setTimeout(() => {
      printWindow.print();
    }, 250);
  };

  return (
    <div className="App">
      <div className="container">
        {/* Header */}
        <h1>Packing Sheet Generator</h1>
        <p className="subtitle">Upload a MESA Excel export and download packing sheets as Word documents</p>

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
              <div className="upload-icon">XLSX</div>
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
                  <p>{files.length} file{files.length > 1 ? 's' : ''} loaded</p>
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
                        x
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
                <div className="mode-toggle" role="group" aria-label="Packing sheet mode">
                  {Object.entries(LIST_MODES).map(([mode, option]) => (
                    <button
                      key={mode}
                      type="button"
                      className={`mode-option ${listMode === mode ? 'active' : ''}`}
                      onClick={() => setListMode(mode)}
                    >
                      {option.label}
                    </button>
                  ))}
                </div>
                <div className="sheet-summary">
                  <span>{selectedPackingSheetCount} of {packingSheets.length} selected</span>
                  <span>{files[selectedFileIndex].sheetName}</span>
                </div>
                <div className="transpose-toggle">
                  <label>
                    <input
                      type="checkbox"
                      checked={includeBorders}
                      onChange={(e) => setIncludeBorders(e.target.checked)}
                    />
                    <span>Show table borders</span>
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
                <div className="item-selection">
                  <div className="item-selection-header">
                    <span>Item pages</span>
                    <div className="selection-actions">
                      <button type="button" onClick={() => setAllPackingSheetsSelected(true)}>All</button>
                      <button type="button" onClick={() => setAllPackingSheetsSelected(false)}>None</button>
                    </div>
                  </div>
                  <div className="item-toggle-list">
                    {packingSheets.map((packingSheet) => (
                      <label key={packingSheet.itemKey} className="item-toggle">
                        <input
                          type="checkbox"
                          checked={isPackingSheetSelected(packingSheet)}
                          onChange={(e) => setPackingSheetSelected(packingSheet, e.target.checked)}
                        />
                        <span className="item-toggle-copy">
                          <span className="item-toggle-name">{packingSheet.itemName}</span>
                          <span className="item-toggle-count">
                            {packingSheet.recipients.length} address{packingSheet.recipients.length === 1 ? '' : 'es'}
                          </span>
                        </span>
                      </label>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Right side - Preview section */}
          {files.length > 0 && files[selectedFileIndex] && showPreview && (
            <div className="right-section">
              <div className="preview-section">
                <h3>{LIST_MODES[listMode].label} Preview - {files[selectedFileIndex].originalName}</h3>
                {/* Document preview styled like Word */}
                <div className="document-preview">
                  {selectedPackingSheets.length > 0 ? (
                    selectedPackingSheets.map((packingSheet) => (
                      <div key={packingSheet.itemKey} className="document-page">
                        <h2 className="document-title">{packingSheet.itemName}</h2>
                        <p className="document-list-mode">{LIST_MODES[listMode].label}</p>
                        <p className="document-count">
                          {packingSheet.recipients.length} address{packingSheet.recipients.length === 1 ? '' : 'es'}
                        </p>
                        <div className="table-container">
                          <table className={`document-table ${includeBorders ? '' : 'no-borders'}`}>
                            <tbody>
                              {packingSheet.rows.map((row, rowIndex) => (
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
                    ))
                  ) : (
                    <div className="document-page">
                      <h2 className="document-title">
                        {packingSheets.length > 0 ? 'No Selected Pages' : 'No Packing Sheets'}
                      </h2>
                      <p className="document-count">
                        {packingSheets.length > 0 ? 'Select an item page to preview it.' : '0 addresses'}
                      </p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Download Buttons */}
        <div className="download-buttons-container">
          <button
            className="download-btn"
            onClick={downloadAllPackingSheets}
            disabled={packingSheets.length === 0 || selectedPackingSheetCount === 0}
          >
            Download Selected Packing Sheets ({selectedPackingSheetCount})
          </button>
          <button
            className="download-btn print-btn"
            onClick={printSelectedPackingSheets}
            disabled={packingSheets.length === 0 || selectedPackingSheetCount === 0}
          >
            Print Selected Packing Sheets ({selectedPackingSheetCount})
          </button>
        </div>
      </div>
    </div>
  );
}

export default App;
