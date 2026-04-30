# CLiCK Packing Sheet Generator

Browser-based packing sheet generator for CLiCK/Mesa spreadsheet exports.

The app reads uploaded `.xlsx` and `.xls` files in the browser, detects the packing sheet data, previews item pages, and generates Word documents for packing workflows. Uploaded workbook data is processed locally in the browser and is not sent to a server by this app.

## Features

- Upload one or more Excel workbooks.
- Upload a folder of Excel files.
- Automatically select the best packing-sheet tab from multi-sheet workbooks.
- Detect address, household-size, and item columns from common header names.
- Generate a single selected item page as `.docx`.
- Generate all item pages as separate `.docx` files in a `.zip`.
- Preview selected pages before generating or printing.
- Print only the selected preview pages.
- Toggle preview page selection with all/none controls.
- Preload a synthetic example workbook for local testing and demos.

## Expected Workbook Format

Workbooks should include one sheet with recipient rows and item columns. Common columns include:

- `Address`
- `Household size`
- Item columns such as `Item Name - D` or `Item Name - DN`

Column suffixes control the page type:

- `-D`: marked values are treated as `DO Want`
- `-DN`: marked values are treated as `DO NOT Want`

The suffix is removed from the printed item name.

Example structure using placeholder data only:

```text
Address              Household size    Diapers - D    Cat Food - DN
[street address]     [number]          [marked]       [blank]
[street address]     [number]          [blank]        [marked]
```

## Privacy Notes

Do not commit real recipient spreadsheets, generated packing sheets, or files containing names, addresses, phone numbers, email addresses, or household details.

This repository should contain only application source code, configuration, documentation with placeholder data, and synthetic demo files. The app preloads `public/ExampleData.xlsx`, so that file is public when the site is deployed.

## Local Development

Install dependencies:

```powershell
npm install
```

Start the local development server:

```powershell
npm start
```

Build the production site:

```powershell
npm run build
```

## GitHub Pages

This app can be hosted on GitHub Pages because it builds to static files.

The included workflow at `.github/workflows/pages.yml` installs dependencies, builds the React app, uploads the build artifact, and deploys it to GitHub Pages.

To enable GitHub Pages:

1. Push the repository to GitHub.
2. Open the repository settings.
3. Go to `Pages`.
4. Set the source to `GitHub Actions`.
5. Push to `main`.

The deployed site is public. Keep private workbook files out of the repository.
