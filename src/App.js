// Import necessary React hooks and libraries
import React, { useEffect, useRef, useState } from 'react';
import * as XLSX from 'xlsx'; // Library for reading Excel files
import { AlignmentType, Document, Packer, Paragraph, TabStopType, TextRun } from 'docx'; // Library for creating Word documents
import { saveAs } from 'file-saver'; // Library for downloading files
import JSZip from 'jszip'; // Library for creating ZIP files
import './App.css';

const EXAMPLE_WORKBOOK_NAME = 'ExampleData.xlsx';
const EXAMPLE_WORKBOOK_PATH = `${process.env.PUBLIC_URL || ''}/${EXAMPLE_WORKBOOK_NAME}`;

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
  auto: {
    label: 'Auto by item',
    title: 'Auto'
  },
  white: {
    label: 'DO want',
    title: 'DO Want'
  },
  black: {
    label: 'DO NOT want',
    title: 'DO NOT Want'
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

const getPackingSheetHeading = (packingSheet) => {
  if (packingSheet.itemName) return packingSheet.itemName;
  return `✅ These addresses ${packingSheet.listLabel}...${packingSheet.itemName}`;
};

const getPackingSheetStatusLine = (packingSheet) => {
  if (packingSheet.listMode === 'black') {
    return '\u274E NO (These addresses DO NOT want) \u274E';
  }

  return '\u2705 YES (These addresses DO want) \u2705';
};

const getPackingSheetStatusColor = (packingSheet) => {
  return packingSheet.listMode === 'black' ? '#C00000' : '#008000';
};

const getHouseholdSizeValue = (householdSize) => {
  return normalizeCell(householdSize) || 'N/A';
};

const PACKING_SHEET_MARGIN_TWIPS = 360;
const QUANTITY_COLUMN_LABEL = 'Quantity';
const RECIPIENT_ADDRESS_TAB_TWIPS = 2000;
const FONT_FAMILIES = ['Arial', 'Calibri', 'Times New Roman', 'Georgia', 'Verdana'];
const LINE_SPACING_OPTIONS = [
  { value: 'single', label: 'Single', multiplier: 1 },
  { value: '1.15', label: '1.15', multiplier: 1.15 },
  { value: '1.5', label: '1.5', multiplier: 1.5 },
  { value: 'double', label: 'Double', multiplier: 2 },
];
const DEFAULT_FONT_SETTINGS = {
  family: 'Arial',
  sizePt: 18,
  color: '#000000',
  bold: true,
  italic: false,
  underline: false,
  lineSpacing: 'double',
};

const clampNumber = (value, min, max) => {
  return Math.min(max, Math.max(min, value));
};

const normalizeFontSettings = (settings) => {
  const merged = { ...DEFAULT_FONT_SETTINGS, ...settings };
  const family = FONT_FAMILIES.includes(merged.family)
    ? merged.family
    : DEFAULT_FONT_SETTINGS.family;
  const color = /^#[0-9a-f]{6}$/i.test(merged.color)
    ? merged.color
    : DEFAULT_FONT_SETTINGS.color;
  const lineSpacing = LINE_SPACING_OPTIONS.some(option => option.value === merged.lineSpacing)
    ? merged.lineSpacing
    : DEFAULT_FONT_SETTINGS.lineSpacing;

  return {
    family,
    sizePt: clampNumber(Number(merged.sizePt) || DEFAULT_FONT_SETTINGS.sizePt, 6, 24),
    color,
    bold: Boolean(merged.bold),
    italic: Boolean(merged.italic),
    underline: Boolean(merged.underline),
    lineSpacing,
  };
};

const getDocxColor = (color) => color.replace('#', '').toUpperCase();

const getPackingSheetTextSizing = (fontSettings) => {
  const normalizedSettings = normalizeFontSettings(fontSettings);
  const lineSpacingOption = LINE_SPACING_OPTIONS.find(option => option.value === normalizedSettings.lineSpacing)
    || LINE_SPACING_OPTIONS[0];
  const recipientFontPt = normalizedSettings.sizePt;
  const headingFontPt = clampNumber(recipientFontPt + 2, 8, 28);
  const recipientGapPt = 0;

  return {
    ...normalizedSettings,
    headingFontPt,
    recipientFontPt,
    recipientGapPt,
    headingSpacingAfterPt: Math.round(clampNumber(headingFontPt * 0.6, 8, 18)),
    recipientSpacingAfterPt: recipientGapPt,
    recipientLineHeight: lineSpacingOption.multiplier,
    docxLineSpacing: Math.round(lineSpacingOption.multiplier * 240),
    docxColor: getDocxColor(normalizedSettings.color),
  };
};

const getPackingSheetPreviewStyle = (textSizing) => {
  return {
    '--document-font-family': textSizing.family,
    '--document-heading-size': `${textSizing.headingFontPt}pt`,
    '--document-heading-space': `${textSizing.headingSpacingAfterPt}pt`,
    '--document-recipient-size': `${textSizing.recipientFontPt}pt`,
    '--document-recipient-gap': `${textSizing.recipientGapPt}pt`,
    '--document-recipient-line-height': String(textSizing.recipientLineHeight),
    '--document-font-color': textSizing.color,
    '--document-font-weight': textSizing.bold ? '700' : '400',
    '--document-font-style': textSizing.italic ? 'italic' : 'normal',
    '--document-font-decoration': textSizing.underline ? 'underline' : 'none',
  };
};

const getPackingSheetPrintStyle = (textSizing) => {
  return [
    `--document-font-family: ${textSizing.family}`,
    `--heading-font-size: ${textSizing.headingFontPt}pt`,
    `--heading-space: ${textSizing.headingSpacingAfterPt}pt`,
    `--recipient-font-size: ${textSizing.recipientFontPt}pt`,
    `--recipient-gap: ${textSizing.recipientGapPt}pt`,
    `--recipient-line-height: ${textSizing.recipientLineHeight}`,
    `--document-font-color: ${textSizing.color}`,
    `--document-font-weight: ${textSizing.bold ? 700 : 400}`,
    `--document-font-style: ${textSizing.italic ? 'italic' : 'normal'}`,
    `--document-font-decoration: ${textSizing.underline ? 'underline' : 'none'}`,
  ].join('; ');
};

const getDocxRunStyle = (textSizing, sizePt) => ({
  size: sizePt * 2,
  font: textSizing.family,
  color: textSizing.docxColor,
  bold: textSizing.bold,
  italics: textSizing.italic,
  ...(textSizing.underline ? { underline: {} } : {}),
});

const cleanItemLabel = (label) => {
  return normalizeCell(label)
    .replace(/\s*-\s*D(?:N)?\s*$/i, '')
    .trim();
};

const getItemModeMarker = (label) => {
  const cleanLabel = normalizeCell(label);
  if (/\s*-\s*DN\s*$/i.test(cleanLabel)) return 'black';
  if (/\s*-\s*D\s*$/i.test(cleanLabel)) return 'white';
  return null;
};

const isGettingItem = (value) => {
  const raw = normalizeCell(value);
  if (!raw) return false;

  const key = normalizeKey(raw).replace(/'/g, '');
  if (!key || NEGATIVE_VALUES.has(key)) return false;
  if (key.startsWith('no ') || key.startsWith('do not ') || key.startsWith('dont ')) return false;

  return true;
};

const isNotGettingItem = (value) => {
  const raw = normalizeCell(value);
  if (!raw) return false;

  const key = normalizeKey(raw).replace(/'/g, '');
  return NEGATIVE_VALUES.has(key)
    || key.startsWith('no ')
    || key.startsWith('do not ')
    || key.startsWith('dont ');
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

  let inferredItemMode = 'white';
  const itemColumns = headers.reduce((columns, header, index) => {
    const sourceLabel = normalizeCell(header);
    if (index < firstItemIndex || index === addressIndex || index === householdSizeIndex) return columns;
    if (!isItemHeader(sourceLabel)) return columns;

    const markedMode = getItemModeMarker(sourceLabel);
    if (markedMode) {
      inferredItemMode = markedMode;
    }

    columns.push({
      index,
      sourceLabel,
      label: cleanItemLabel(sourceLabel),
      listMode: markedMode || inferredItemMode
    });
    return columns;
  }, []);

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
    const effectiveListMode = listMode === 'auto' ? itemColumn.listMode : listMode;
    const listLabel = LIST_MODES[effectiveListMode].title;
    const useExplicitOptOuts = listMode === 'auto' && effectiveListMode === 'black';
    const recipients = rows
      .map(row => {
        const address = normalizeCell(row[config.addressIndex]);
        const householdSize = config.householdSizeIndex >= 0
          ? normalizeCell(row[config.householdSizeIndex])
          : '';
        const itemValue = normalizeCell(row[itemColumn.index]);
        const gettingItem = isGettingItem(itemValue);
        const notGettingItem = isNotGettingItem(itemValue);

        return {
          address,
          householdSize,
          itemValue,
          gettingItem,
          notGettingItem
        };
      })
      .filter(recipient => recipient.address !== '')
      .filter(recipient => {
        if (effectiveListMode === 'white') return recipient.gettingItem;
        if (useExplicitOptOuts) return recipient.notGettingItem;
        return !recipient.gettingItem;
      });

    return {
      itemKey: `${itemColumn.index}-${itemColumn.sourceLabel}`,
      itemName: itemColumn.label,
      title: `${itemColumn.label} - ${listLabel}`,
      listMode: effectiveListMode,
      listLabel,
      recipients
    };
  });
};

const parseWorkbookFile = async (file, index = 0) => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  const extracted = extractBestWorksheet(workbook);

  if (!extracted) return null;

  return {
    id: `${Date.now()}-${index}-${file.name}`,
    name: sanitizeFilePart(file.name.replace(/\.[^/.]+$/, '')),
    data: extracted.data,
    originalName: workbook.SheetNames.length > 1
      ? `${file.name} - ${extracted.sheetName}`
      : file.name,
    sheetName: extracted.sheetName,
    transposeByDefault: extracted.transposeByDefault,
    splitByDefault: extracted.splitByDefault,
    packingConfig: extracted.packingConfig
  };
};

function App() {
  // State management
  const [files, setFiles] = useState([]); // Array of uploaded Excel files with their data
  const [selectedFileIndex, setSelectedFileIndex] = useState(0); // Index of currently selected file for preview
  const [isDragging, setIsDragging] = useState(false); // Track drag-and-drop state for visual feedback
  const listMode = 'auto'; // Packing sheet type is determined by each item column suffix.
  const [showPreview, setShowPreview] = useState(true); // Whether to show the document preview
  const [fontSettingsOpen, setFontSettingsOpen] = useState(false);
  const [fontSettings, setFontSettings] = useState(DEFAULT_FONT_SETTINGS);
  const [selectedPackingSheetKeys, setSelectedPackingSheetKeys] = useState({}); // Per item/page output selection
  const [selectedItemKey, setSelectedItemKey] = useState(''); // Current item used by single-sheet actions
  const previewContainerRef = useRef(null);
  const previewPageRefs = useRef({});
  const previewScrollLockRef = useRef(false);
  const previewScrollLockTimeoutRef = useRef(null);

  useEffect(() => {
    let isMounted = true;

    const loadExampleWorkbook = async () => {
      try {
        const response = await fetch(EXAMPLE_WORKBOOK_PATH);
        if (!response.ok) return;

        const blob = await response.blob();
        const file = new File([blob], EXAMPLE_WORKBOOK_NAME, {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        const exampleFile = await parseWorkbookFile(file);

        if (!isMounted || !exampleFile) return;

        setFiles(prev => (prev.length > 0 ? prev : [exampleFile]));
        setSelectedFileIndex(0);
      } catch (error) {
        console.warn('Unable to load example workbook:', error);
      }
    };

    loadExampleWorkbook();

    return () => {
      isMounted = false;
    };
  }, []);

  useEffect(() => {
    return () => {
      if (previewScrollLockTimeoutRef.current) {
        clearTimeout(previewScrollLockTimeoutRef.current);
      }
    };
  }, []);

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
          // Newer MESA exports can include title/metadata rows or a summary
          // sheet before the actual packing data.
          const parsedFile = await parseWorkbookFile(file, i);
          if (parsedFile) newFiles.push(parsedFile);
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

  const updateFontSetting = (key, value) => {
    setFontSettings(prev => normalizeFontSettings({
      ...prev,
      [key]: value,
    }));
  };

  const resetFontSettings = () => {
    setFontSettings(DEFAULT_FONT_SETTINGS);
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
   * Creates a Word document (.docx) from one packing sheet.
   * @param {Object} packingSheet - Packing sheet to include
   * @returns {Blob} Word document as a blob
   */
  const createPackingSheetDocument = async (packingSheet) => {
    const textSizing = getPackingSheetTextSizing(fontSettings);
    const recipientRunStyle = getDocxRunStyle(textSizing, textSizing.recipientFontPt);
    const headingRunStyle = getDocxRunStyle(textSizing, textSizing.headingFontPt);
    const statusRunStyle = {
      ...headingRunStyle,
      color: getDocxColor(getPackingSheetStatusColor(packingSheet)),
    };
    const columnHeaderRunStyle = { ...recipientRunStyle, bold: true };
    const recipientParagraphs = packingSheet.recipients.length > 0
      ? packingSheet.recipients.map(recipient => new Paragraph({
        wordWrap: false,
        tabStops: [
          {
            type: TabStopType.LEFT,
            position: RECIPIENT_ADDRESS_TAB_TWIPS,
          },
        ],
        children: [
          new TextRun({
            text: getHouseholdSizeValue(recipient.householdSize),
            ...recipientRunStyle,
          }),
          new TextRun({
            text: '\t',
            ...recipientRunStyle,
          }),
          new TextRun({
            text: recipient.address,
            ...recipientRunStyle,
          }),
        ],
        spacing: {
          line: textSizing.docxLineSpacing,
          lineRule: 'auto',
          after: textSizing.recipientSpacingAfterPt * 20,
        },
      }))
      : [
        new Paragraph({
          wordWrap: false,
          children: [
            new TextRun({
              text: 'No addresses marked for this item.',
              italics: true,
              ...recipientRunStyle,
            }),
          ],
          spacing: {
            line: textSizing.docxLineSpacing,
            lineRule: 'auto',
          },
        }),
      ];

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: PACKING_SHEET_MARGIN_TWIPS,
                right: PACKING_SHEET_MARGIN_TWIPS,
                bottom: PACKING_SHEET_MARGIN_TWIPS,
                left: PACKING_SHEET_MARGIN_TWIPS,
              },
            },
          },
          children: [
            new Paragraph({
              wordWrap: false,
              children: [
                new TextRun({
                  text: getPackingSheetHeading(packingSheet),
                  ...headingRunStyle,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                after: 80,
              },
            }),
            new Paragraph({
              wordWrap: false,
              children: [
                new TextRun({
                  text: getPackingSheetStatusLine(packingSheet),
                  ...statusRunStyle,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                after: textSizing.headingSpacingAfterPt * 20,
              },
            }),
            new Paragraph({
              wordWrap: false,
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: RECIPIENT_ADDRESS_TAB_TWIPS,
                },
              ],
              children: [
                new TextRun({
                  text: QUANTITY_COLUMN_LABEL,
                  ...columnHeaderRunStyle,
                }),
                new TextRun({
                  text: '\t',
                  ...columnHeaderRunStyle,
                }),
                new TextRun({
                  text: 'Address',
                  ...columnHeaderRunStyle,
                }),
              ],
              spacing: {
                after: 80,
              },
            }),
            ...recipientParagraphs,
          ],
        },
      ],
    });

    // Convert document to blob for download
    return await Packer.toBlob(doc);
  };

  const currentFile = files[selectedFileIndex];
  const packingSheets = getPackingSheets(currentFile, listMode);
  const packingSheetTextSizing = getPackingSheetTextSizing(fontSettings);
  const getSelectionKey = (packingSheet) => {
    const fileKey = currentFile?.id || currentFile?.originalName || 'current-file';
    return `${fileKey}:${listMode}:${packingSheet.itemKey || packingSheet.itemName}`;
  };
  const isPackingSheetSelected = (packingSheet) => selectedPackingSheetKeys[getSelectionKey(packingSheet)] !== false;
  const selectedPackingSheets = packingSheets.filter(isPackingSheetSelected);
  const selectedPackingSheetCount = selectedPackingSheets.length;
  const selectedItemPackingSheet = packingSheets.find(packingSheet => packingSheet.itemKey === selectedItemKey)
    || packingSheets[0]
    || null;
  const selectedPreviewItemKey = selectedItemPackingSheet?.itemKey || '';

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

  const lockPreviewScrollSync = () => {
    previewScrollLockRef.current = true;

    if (previewScrollLockTimeoutRef.current) {
      clearTimeout(previewScrollLockTimeoutRef.current);
    }

    previewScrollLockTimeoutRef.current = setTimeout(() => {
      previewScrollLockRef.current = false;
      previewScrollLockTimeoutRef.current = null;
    }, 300);
  };

  const selectSinglePackingSheet = (itemKey) => {
    const packingSheet = packingSheets.find(sheet => sheet.itemKey === itemKey);
    lockPreviewScrollSync();
    setSelectedItemKey(itemKey);

    if (packingSheet) {
      setPackingSheetSelected(packingSheet, true);
    }
  };

  const syncSelectedItemToPreviewScroll = () => {
    if (previewScrollLockRef.current) return;

    const container = previewContainerRef.current;
    if (!container || selectedPackingSheets.length === 0) return;

    const containerBounds = container.getBoundingClientRect();
    const containerCenter = containerBounds.left + (containerBounds.width / 2);
    let nearestPackingSheet = null;
    let nearestDistance = Infinity;

    selectedPackingSheets.forEach(packingSheet => {
      const page = previewPageRefs.current[packingSheet.itemKey];
      if (!page) return;

      const pageBounds = page.getBoundingClientRect();
      const pageCenter = pageBounds.left + (pageBounds.width / 2);
      const distance = Math.abs(pageCenter - containerCenter);

      if (distance < nearestDistance) {
        nearestDistance = distance;
        nearestPackingSheet = packingSheet;
      }
    });

    if (nearestPackingSheet && nearestPackingSheet.itemKey !== selectedItemPackingSheet?.itemKey) {
      setSelectedItemKey(nearestPackingSheet.itemKey);
    }
  };

  useEffect(() => {
    if (packingSheets.length === 0) {
      if (selectedItemKey !== '') setSelectedItemKey('');
      return;
    }

    if (!packingSheets.some(packingSheet => packingSheet.itemKey === selectedItemKey)) {
      setSelectedItemKey(packingSheets[0].itemKey);
    }
  }, [packingSheets, selectedItemKey]);

  useEffect(() => {
    if (!showPreview || !selectedPreviewItemKey) return;

    const page = previewPageRefs.current[selectedPreviewItemKey];
    if (!page) return;

    lockPreviewScrollSync();
    page.scrollIntoView({
      behavior: 'auto',
      block: 'nearest',
      inline: 'center',
    });
  }, [selectedPreviewItemKey, selectedPackingSheetCount, showPreview]);

  const generatePackingSheetFiles = async (sheetsToGenerate, fileLabel) => {
    try {
      if (sheetsToGenerate.length === 1) {
        const packingSheet = sheetsToGenerate[0];
        const blob = await createPackingSheetDocument(packingSheet);
        saveAs(blob, `${sanitizeFilePart(packingSheet.itemName)}_${packingSheet.listMode}.docx`);
        alert(`Successfully created ${packingSheet.itemName} packing sheet!`);
        return;
      }

      const zip = new JSZip();
      const folder = zip.folder(`${sanitizeFilePart(currentFile.name)}_${fileLabel}_packing_sheets`);

      for (const packingSheet of sheetsToGenerate) {
        const blob = await createPackingSheetDocument(packingSheet);
        folder.file(`${sanitizeFilePart(packingSheet.itemName)}_${packingSheet.listMode}.docx`, blob);
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      saveAs(zipBlob, `${sanitizeFilePart(currentFile.name)}_${fileLabel}_packing_sheets.zip`);
      alert(`Successfully created ${sheetsToGenerate.length} packing sheets!`);
    } catch (error) {
      console.error('Error creating packing sheets:', error);
      alert('Error creating packing sheets. Please try again.');
    }
  };

  const downloadCurrentPackingSheet = async () => {
    if (!currentFile || packingSheets.length === 0) {
      alert('Please upload a MESA Excel file with address and item columns first!');
      return;
    }

    if (!selectedItemPackingSheet) {
      alert('Please select an item page to download.');
      return;
    }

    await generatePackingSheetFiles([selectedItemPackingSheet], 'current');
  };

  const downloadSelectedPackingSheets = async () => {
    if (!currentFile || packingSheets.length === 0) {
      alert('Please upload a MESA Excel file with address and item columns first!');
      return;
    }

    if (selectedPackingSheetCount === 0) {
      alert('Please select at least one item page to download.');
      return;
    }

    await generatePackingSheetFiles(selectedPackingSheets, 'selected');
  };

  const downloadAllPackingSheets = async () => {
    if (!currentFile || packingSheets.length === 0) {
      alert('Please upload a MESA Excel file with address and item columns first!');
      return;
    }

    await generatePackingSheetFiles(packingSheets, 'all');
  };

  const printPackingSheets = (sheetsToPrint, emptySelectionMessage) => {
    if (!currentFile || packingSheets.length === 0) {
      alert('Please upload a MESA Excel file with address and item columns first!');
      return;
    }

    if (sheetsToPrint.length === 0) {
      alert(emptySelectionMessage);
      return;
    }

    const pageStyle = getPackingSheetPrintStyle(packingSheetTextSizing);
    const pagesHtml = sheetsToPrint.map(packingSheet => `
      <section class="page" style="${pageStyle}">
        <h1>${escapeHtml(getPackingSheetHeading(packingSheet))}</h1>
        <p class="document-status" style="--document-status-color: ${getPackingSheetStatusColor(packingSheet)}">${escapeHtml(getPackingSheetStatusLine(packingSheet))}</p>
        <div class="recipient-line recipient-header">
          <span>${QUANTITY_COLUMN_LABEL}</span>
          <span>Address</span>
        </div>
        <div class="recipient-list">
          ${packingSheet.recipients.length > 0
            ? packingSheet.recipients.map(recipient => `
              <p class="recipient-line">
                <span>${escapeHtml(getHouseholdSizeValue(recipient.householdSize))}</span>
                <span>${escapeHtml(recipient.address)}</span>
              </p>
            `).join('')
            : '<p class="recipient-line muted">No addresses marked for this item.</p>'}
        </div>
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
              padding: 0.25in;
              page-break-after: always;
              font-family: var(--document-font-family, Arial), Helvetica, sans-serif;
              color: var(--document-font-color, #000);
            }
            .page:last-child {
              page-break-after: auto;
            }
            h1 {
              margin: 0 0 4pt 0;
              text-align: center;
              font-size: var(--heading-font-size, 24pt);
              line-height: 1.15;
              font-weight: var(--document-font-weight, 400);
              font-style: var(--document-font-style, normal);
              text-decoration: var(--document-font-decoration, none);
            }
            .document-status {
              margin: 0 0 var(--heading-space, 24pt) 0;
              color: var(--document-status-color, var(--document-font-color, #000));
              text-align: center;
              font-size: var(--heading-font-size, 24pt);
              line-height: 1.15;
              font-weight: var(--document-font-weight, 400);
              font-style: var(--document-font-style, normal);
              text-decoration: var(--document-font-decoration, none);
            }
            .recipient-list {
              display: grid;
              gap: var(--recipient-gap, 8pt);
            }
            .recipient-line {
              display: grid;
              grid-template-columns: 5.2em minmax(0, 1fr);
              gap: 18px;
              margin: 0;
              font-size: var(--recipient-font-size, 16pt);
              line-height: var(--recipient-line-height, 1.18);
              white-space: nowrap;
              font-weight: var(--document-font-weight, 400);
              font-style: var(--document-font-style, normal);
              text-decoration: var(--document-font-decoration, none);
            }
            .recipient-header {
              margin-bottom: 6pt;
              font-weight: 700;
            }
            .muted {
              color: #666;
              font-style: italic;
            }
            @page {
              margin: 0.25in;
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

  const printCurrentPackingSheet = () => {
    printPackingSheets(
      selectedItemPackingSheet ? [selectedItemPackingSheet] : [],
      'Please select an item page to print.'
    );
  };

  const printSelectedPackingSheets = () => {
    printPackingSheets(selectedPackingSheets, 'Please select at least one item page to print.');
  };

  const printAllPackingSheets = () => {
    printPackingSheets(packingSheets, 'Please upload a MESA Excel file with address and item columns first!');
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
                <div className="sheet-summary">
                  <span>{selectedPackingSheetCount} of {packingSheets.length} selected</span>
                  <span>{files[selectedFileIndex].sheetName}</span>
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
                <div className={`font-settings ${fontSettingsOpen ? 'open' : ''}`}>
                  <div className="font-settings-header">
                    <span>Font</span>
                    <button
                      type="button"
                      className="settings-icon-btn"
                      aria-label="Font settings"
                      title="Font settings"
                      onClick={() => setFontSettingsOpen(prev => !prev)}
                    >
                      <span aria-hidden="true">&#9881;</span>
                    </button>
                  </div>
                  {fontSettingsOpen && (
                    <div className="font-settings-panel">
                      <label className="font-control font-size-control">
                        <span>Size</span>
                        <input
                          type="range"
                          min="6"
                          max="24"
                          step="1"
                          value={fontSettings.sizePt}
                          onChange={(e) => updateFontSetting('sizePt', e.target.value)}
                        />
                        <input
                          type="number"
                          min="6"
                          max="24"
                          value={fontSettings.sizePt}
                          onChange={(e) => updateFontSetting('sizePt', e.target.value)}
                        />
                      </label>
                      <label className="font-control">
                        <span>Family</span>
                        <select
                          value={fontSettings.family}
                          onChange={(e) => updateFontSetting('family', e.target.value)}
                        >
                          {FONT_FAMILIES.map(fontFamily => (
                            <option key={fontFamily} value={fontFamily}>{fontFamily}</option>
                          ))}
                        </select>
                      </label>
                      <label className="font-control">
                        <span>Line spacing</span>
                        <select
                          value={fontSettings.lineSpacing}
                          onChange={(e) => updateFontSetting('lineSpacing', e.target.value)}
                        >
                          {LINE_SPACING_OPTIONS.map(option => (
                            <option key={option.value} value={option.value}>{option.label}</option>
                          ))}
                        </select>
                      </label>
                      <div className="format-controls">
                        <button
                          type="button"
                          className={`format-toggle ${fontSettings.bold ? 'active' : ''}`}
                          aria-pressed={fontSettings.bold}
                          onClick={() => updateFontSetting('bold', !fontSettings.bold)}
                        >
                          B
                        </button>
                        <button
                          type="button"
                          className={`format-toggle italic-toggle ${fontSettings.italic ? 'active' : ''}`}
                          aria-pressed={fontSettings.italic}
                          onClick={() => updateFontSetting('italic', !fontSettings.italic)}
                        >
                          I
                        </button>
                        <button
                          type="button"
                          className={`format-toggle underline-toggle ${fontSettings.underline ? 'active' : ''}`}
                          aria-pressed={fontSettings.underline}
                          onClick={() => updateFontSetting('underline', !fontSettings.underline)}
                        >
                          U
                        </button>
                        <label className="color-control" title="Font color">
                          <span>Color</span>
                          <input
                            type="color"
                            value={fontSettings.color}
                            onChange={(e) => updateFontSetting('color', e.target.value)}
                          />
                        </label>
                      </div>
                      <button type="button" className="reset-font-btn" onClick={resetFontSettings}>
                        Reset font
                      </button>
                    </div>
                  )}
                </div>
                <div className="item-selection">
                  <div className="item-selection-header">
                    <span>Generate one item</span>
                  </div>
                  <label className="single-item-select" htmlFor="single-item-select">
                    <span>Selected item</span>
                    <select
                      id="single-item-select"
                      value={selectedItemPackingSheet?.itemKey || ''}
                      onChange={(e) => selectSinglePackingSheet(e.target.value)}
                    >
                      {packingSheets.map((packingSheet) => (
                        <option key={packingSheet.itemKey} value={packingSheet.itemKey}>
                          {packingSheet.itemName} - {packingSheet.listLabel}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>
                <div className="item-selection">
                  <div className="item-selection-header">
                    <span>Preview / print pages</span>
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
                            {packingSheet.listLabel} - {packingSheet.recipients.length} address{packingSheet.recipients.length === 1 ? '' : 'es'}
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
                <h3>Packing Sheet Preview - {files[selectedFileIndex].originalName}</h3>
                {/* Document preview styled like Word */}
                <div
                  className="document-preview"
                  ref={previewContainerRef}
                  onScroll={syncSelectedItemToPreviewScroll}
                >
                  {selectedPackingSheets.length > 0 ? (
                    selectedPackingSheets.map((packingSheet) => {
                      const previewStyle = getPackingSheetPreviewStyle(packingSheetTextSizing);

                      return (
                        <div
                          key={packingSheet.itemKey}
                          className="document-page"
                          style={previewStyle}
                          ref={(page) => {
                            if (page) {
                              previewPageRefs.current[packingSheet.itemKey] = page;
                            } else {
                              delete previewPageRefs.current[packingSheet.itemKey];
                            }
                          }}
                        >
                          <h2 className="document-title">{getPackingSheetHeading(packingSheet)}</h2>
                          <p
                            className="document-status"
                            style={{ '--document-status-color': getPackingSheetStatusColor(packingSheet) }}
                          >
                            {getPackingSheetStatusLine(packingSheet)}
                          </p>
                          <div className="recipient-line recipient-header">
                            <span>{QUANTITY_COLUMN_LABEL}</span>
                            <span>Address</span>
                          </div>
                          <div className="recipient-list">
                            {packingSheet.recipients.length > 0 ? (
                              packingSheet.recipients.map((recipient, recipientIndex) => (
                                <p key={recipientIndex} className="recipient-line">
                                  <span>{getHouseholdSizeValue(recipient.householdSize)}</span>
                                  <span>{recipient.address}</span>
                                </p>
                              ))
                            ) : (
                              <p className="recipient-line muted">No addresses marked for this item.</p>
                            )}
                          </div>
                        </div>
                      );
                    })
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

        {/* Packing sheet actions */}
        <div className="download-buttons-container">
          <div className="packing-action-column">
            <button
              type="button"
              className="download-btn"
              onClick={downloadCurrentPackingSheet}
              disabled={packingSheets.length === 0 || !selectedItemPackingSheet}
            >
              Download Current Packing Sheet
            </button>
            <button
              type="button"
              className="download-btn download-selected-btn"
              onClick={downloadSelectedPackingSheets}
              disabled={packingSheets.length === 0 || selectedPackingSheetCount === 0}
            >
              Download Selected Packing Sheets
            </button>
            <button
              type="button"
              className="download-btn download-all-btn"
              onClick={downloadAllPackingSheets}
              disabled={packingSheets.length === 0}
            >
              Download All Packing Sheets
            </button>
          </div>
          <div className="packing-action-column">
            <button
              type="button"
              className="download-btn print-current-btn"
              onClick={printCurrentPackingSheet}
              disabled={packingSheets.length === 0 || !selectedItemPackingSheet}
            >
              Print Current Packing Sheet
            </button>
            <button
              type="button"
              className="download-btn print-selected-btn"
              onClick={printSelectedPackingSheets}
              disabled={packingSheets.length === 0 || selectedPackingSheetCount === 0}
            >
              Print Selected Packing Sheets
            </button>
            <button
              type="button"
              className="download-btn print-all-btn"
              onClick={printAllPackingSheets}
              disabled={packingSheets.length === 0}
            >
              Print All Packing Sheets
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
