require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const https = require('https');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const GROQ_API_KEY = process.env.GROQ_API_KEY;

const SYSTEM_PROMPT = `You are an AI assistant that controls a spreadsheet. You receive natural language commands and return JSON actions.

Respond ONLY with a JSON array of actions. No explanations, no markdown, just the JSON array.

Supported actions:

1. {"action": "setCell", "row": 0, "col": 0, "value": "Hello"} - Set a cell value (row/col are 0-indexed)
2. {"action": "setCells", "startRow": 0, "startCol": 0, "values": [["A1","B1"],["A2","B2"]]} - Set multiple cells from a 2D array
3. {"action": "addColumn", "index": 0, "name": "Column Name"} - Add a column at index with header name
4. {"action": "addRow", "index": 0, "values": ["val1","val2"]} - Add a row at index with values
5. {"action": "deleteColumn", "index": 0} - Delete a column by index
6. {"action": "deleteRow", "index": 0} - Delete a row by index
7. {"action": "setCellStyle", "row": 0, "col": 0, "style": {"bg": "#ff0000", "color": "#ffffff", "bold": true, "fontSize": 14}} - Style a cell
8. {"action": "setRangeStyle", "startRow": 0, "startCol": 0, "endRow": 2, "endCol": 3, "style": {"bg": "#ffff00"}} - Style a range
9. {"action": "mergeCells", "startRow": 0, "startCol": 0, "endRow": 1, "endCol": 2} - Merge cells
10. {"action": "setColumnWidth", "col": 0, "width": 200} - Set column width in pixels
11. {"action": "setRowHeight", "row": 0, "height": 40} - Set row height in pixels
12. {"action": "setFormula", "row": 0, "col": 0, "formula": "=SUM(A1:A10)"} - Set a formula in a cell
13. {"action": "freezeRow", "count": 1} - Freeze top N rows
14. {"action": "freezeCol", "count": 1} - Freeze left N columns
15. {"action": "addSheet", "name": "Sheet2"} - Add a new sheet
16. {"action": "renameSheet", "oldName": "Sheet1", "newName": "Data"} - Rename a sheet
17. {"action": "clearRange", "startRow": 0, "startCol": 0, "endRow": 5, "endCol": 5} - Clear a range
18. {"action": "setTextAlign", "row": 0, "col": 0, "align": "center"} - Set text alignment (left/center/right)
19. {"action": "setBorder", "startRow": 0, "startCol": 0, "endRow": 2, "endCol": 2, "border": "all"} - Add borders (all/outer/inner/top/bottom/left/right)
20. {"action": "autoFill", "startRow": 0, "startCol": 0, "endRow": 0, "endCol": 0, "fillRows": 10} - Auto-fill a pattern down
21. {"action": "fillRandom", "startRow": 1, "col": 0, "count": 10, "min": 100, "max": 9999} - Fill a column with random numbers
22. {"action": "fillRandom", "startRow": 1, "col": 0, "count": 10, "type": "names"} - Fill with random names
23. {"action": "fillRandom", "startRow": 1, "col": 0, "count": 10, "type": "amounts"} - Fill with random money amounts

Color names supported: red, blue, green, yellow, orange, purple, pink, black, white, gray, lightblue, lightgreen, lightyellow, lightgray.

Be smart about interpreting commands. For example:
- "add random values to income" -> Find the Income column and fill rows 1-10 with random amounts: [{"action": "fillRandom", "startRow": 1, "col": 0, "count": 10, "type": "amounts"}]
- "add a column called Name" -> [{"action": "addColumn", "index": -1, "name": "Name"}]
- "make the header row blue with white text" -> [{"action": "setRangeStyle", "startRow": 0, "startCol": 0, "endRow": 0, "endCol": -1, "style": {"bg": "#0066cc", "color": "#ffffff", "bold": true}}]
- "create a budget sheet with Income, Expense, Balance columns" -> [{"action": "addColumn", "index": 0, "name": "Income"}, {"action": "addColumn", "index": 1, "name": "Expense"}, {"action": "addColumn", "index": 2, "name": "Balance"}]
- If endCol or endRow is -1, it means "to the last used column/row"
- If index is -1 for addColumn, add at the end
- For "fillRandom" type "amounts" generates values like 1500.00, type "names" generates random person names, type "emails" generates random emails, default min/max generates integers`;

function groqRequest(payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const options = {
      hostname: 'api.groq.com',
      path: '/openai/v1/chat/completions',
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${GROQ_API_KEY}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error('Invalid JSON from Groq: ' + data.substring(0, 200))); }
      });
    });

    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

app.post('/api/ai', async (req, res) => {
  const { message, sheetData } = req.body;

  if (!GROQ_API_KEY || GROQ_API_KEY === 'your_groq_api_key_here') {
    return res.json({
      success: false,
      error: 'Please set your GROQ_API_KEY in the .env file. Get a free key at https://console.groq.com'
    });
  }

  try {
    const contextInfo = sheetData ? `\n\nCurrent sheet info: ${JSON.stringify(sheetData)}` : '';

    const data = await groqRequest({
      model: 'llama-3.3-70b-versatile',
      messages: [
        { role: 'system', content: SYSTEM_PROMPT + contextInfo },
        { role: 'user', content: message }
      ],
      temperature: 0.1,
      max_tokens: 2048
    });

    if (data.error) {
      return res.json({ success: false, error: data.error.message });
    }

    let content = data.choices[0].message.content.trim();

    // Strip markdown code blocks if present
    if (content.startsWith('```')) {
      content = content.replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();
    }

    let actions;
    try {
      actions = JSON.parse(content);
      if (!Array.isArray(actions)) actions = [actions];
    } catch (e) {
      console.error('Parse failed. Raw AI response:', content);
      return res.json({ success: false, error: 'AI returned: ' + content.substring(0, 300) });
    }

    res.json({ success: true, actions });
  } catch (err) {
    console.error('AI Error:', err);
    res.json({ success: false, error: err.message });
  }
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n  AI Sheets running at http://localhost:${PORT}\n`);
});
