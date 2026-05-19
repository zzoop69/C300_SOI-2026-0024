# Autonomous Supplier Payment Process

A simple full-stack demo website for checking whether supplier purchase orders, sales orders, and invoices match before payment approval.

## Features

- Dashboard totals for purchase orders, sales orders, checked invoices, matched, mismatched, and pending records
- Upload form for PO, SO, and invoice files
- Beginner-friendly simulated document extraction using form fields
- Document records page
- Matching results page with mismatched fields highlighted
- Automated summary report table
- JSON file storage with dummy data for demo use

## Project Structure

```text
controllers/              Express routes
models/                   JSON storage and matching logic
public/css/               Website styling
views/                    EJS pages
data/                     Created automatically for JSON records
uploads/                  Created automatically for uploaded files
app.js                    Main Express server
```

## How To Run In VS Code Terminal

1. Open this project folder in VS Code.
2. Open the terminal.
3. Install dependencies:

```bash
npm install
```

4. Start the website:

```bash
npm start
```

5. Open this address in your browser:

```text
http://localhost:3000
```

The app creates demo data automatically the first time it runs.
