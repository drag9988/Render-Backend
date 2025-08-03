# PDF Password Protection Feature

## New Endpoint: `/add-password-to-pdf`

### Description
Adds password protection to PDF files using multiple fallback methods for maximum reliability.

### Rate Limiting
- **10 requests per day** per IP address

### Request Format
```
POST /add-password-to-pdf
Content-Type: multipart/form-data

Form fields:
- file: PDF file (max 50MB)
- password: String (4-128 characters)
```

### Response
- **Success (200)**: Returns password-protected PDF file
- **Error (400)**: Invalid file or password
- **Error (408)**: Processing timeout
- **Error (503)**: Service unavailable

### Password Requirements
- Minimum 4 characters
- Maximum 128 characters
- Cannot be empty or whitespace only

### Security Features
- File validation (MIME type, extension, size, header verification)
- Filename sanitization
- Malicious content detection
- Multiple encryption methods:
  1. **qpdf** (primary method - 256-bit AES encryption)
  2. **pdftk** (fallback method)
  3. **Python helper script** (last resort)

### Usage Examples

#### cURL
```bash
curl -X POST http://localhost:3000/add-password-to-pdf \
  -F "file=@document.pdf" \
  -F "password=mySecurePassword123" \
  --output protected_document.pdf
```

#### JavaScript (Fetch API)
```javascript
const formData = new FormData();
formData.append('file', pdfFile);
formData.append('password', 'mySecurePassword123');

const response = await fetch('/add-password-to-pdf', {
  method: 'POST',
  body: formData
});

if (response.ok) {
  const protectedPdf = await response.blob();
  // Handle the protected PDF
} else {
  const error = await response.json();
  console.error('Error:', error.message);
}
```

### Error Handling

#### Common Error Responses
```json
{
  "error": "PDF validation failed",
  "details": ["Invalid file extension", "File too large"]
}
```

```json
{
  "error": "Invalid input",
  "message": "Password must be at least 4 characters long"
}
```

```json
{
  "error": "Password protection timeout",
  "message": "The PDF password protection is taking too long. Please try with a smaller PDF."
}
```

### Technical Implementation
- Uses **qpdf** for 256-bit AES encryption (industry standard)
- Falls back to **pdftk** if qpdf is unavailable
- Final fallback uses Python script with multiple methods
- Comprehensive file validation and sanitization
- Temporary file cleanup for security
- Detailed logging for debugging

### System Requirements
The following tools are installed in the Docker container:
- `qpdf` - Primary PDF encryption tool
- `pdftk` - Fallback PDF manipulation tool
- `python3` - For helper scripts
- `libreoffice` - Document processing

### Similar to Aadhaar PDF Protection
This feature implements password protection similar to Aadhaar PDFs:
- Strong encryption (256-bit AES when using qpdf)
- Document open password protection
- Maintains PDF integrity and formatting
- Compatible with standard PDF readers
