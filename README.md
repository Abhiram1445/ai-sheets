# AI Sheets

An AI-powered spreadsheet that lets you control everything with natural language. Uses Groq API (free) with Llama 3.3.

## Features

- **Google Sheets-like UI** with Luckysheet
- **AI Assistant** - type natural language commands to manipulate the spreadsheet
- **Auto-save** to localStorage
- **Download as .xlsx** using SheetJS
- **20+ actions** including add/delete columns/rows, styling, formulas, random data, borders, etc.

## Setup

```bash
npm install
```

Create a `.env` file:
```
GROQ_API_KEY=your_groq_api_key_here
PORT=3000
```

Get a free API key at https://console.groq.com

## Run

```bash
npm start
```

Open http://localhost:3000

## AI Commands Examples

- "Add columns: Name, Email, Phone"
- "Make header row dark blue with white bold text"
- "Fill column A with random names"
- "Add SUM formulas in row 11"
- "Style alternating rows with light gray"
- "Add borders to the table"
- "Create a budget sheet with Income, Expense, Balance"
