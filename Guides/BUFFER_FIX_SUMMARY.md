# Buffer Compatibility Fix

## Problem
The hybrid parser worked in the test script but failed when called from the Electron app's IPC handler.

## Root Cause
**Different Buffer Types:**
- **Test script**: Uses `fs.readFileSync()` → Returns Node.js **Buffer**
- **Electron IPC**: Uses `FileReader.readAsArrayBuffer()` → Returns **ArrayBuffer**
- **pdf2json**: Expects Node.js **Buffer**, not ArrayBuffer

## Solution
Added automatic buffer conversion in both `hybrid-parser.js` and `utilities.js`:

```javascript
// Convert ArrayBuffer to Buffer if needed (from Electron IPC)
let buffer = fileBuffer;
if (fileBuffer instanceof ArrayBuffer) {
  buffer = Buffer.from(fileBuffer);
} else if (fileBuffer.buffer instanceof ArrayBuffer) {
  // Handle Uint8Array
  buffer = Buffer.from(fileBuffer.buffer);
}
```

## Files Updated

### 1. hybrid-parser.js
- Added buffer conversion in `extractStructuredContent()` method
- Now handles both ArrayBuffer (from Electron) and Buffer (from Node.js)

### 2. utilities.js
- Replaced `pdfjs-dist` with `pdf2json` (fixes DOMMatrix error)
- Added buffer conversion in `extractTextFromPDF()` method
- Now compatible with Electron IPC

### 3. main.js
- Added debug logging to verify buffer types
- Shows buffer type and length for troubleshooting

## Testing

### Test from command line (Node.js Buffer):
```bash
node test-hybrid-parser.js /path/to/file.pdf
```

### Test from Electron app (ArrayBuffer):
1. Run `npm start`
2. Select a PDF file
3. Click Convert
4. Check console for debug output

Both should now work correctly!

## Debug Output
When running the Electron app, you'll see:
```
=== PDF PARSE DEBUG ===
Buffer type: ArrayBuffer
Is ArrayBuffer: true
Is Buffer: false
Buffer length: 23866
=======================
```

After conversion, pdf2json receives a proper Node.js Buffer and parses successfully.
