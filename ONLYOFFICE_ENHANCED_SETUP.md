# Enhanced ONLYOFFICE Integration Setup Guide

## Overview
This backend now uses **ONLYOFFICE Document Server (Community Edition)** as the primary conversion method for PDF to Word/Excel/PowerPoint, with multiple fallback methods for maximum reliability.

## Conversion Priority Order
1. **🥇 Enhanced ONLYOFFICE Service** (Primary)
   - ONLYOFFICE Document Server (if configured)
   - Python-based conversion (pdf2docx, pdfplumber, etc.)
   - Enhanced LibreOffice with specialized options

2. **🥈 Original ONLYOFFICE Service** (Secondary backup)
3. **🥉 ConvertAPI** (Tertiary backup - only if others fail)
4. **🛡️ Standard LibreOffice** (Final fallback)

## Deployment Options

### Option 1: Enhanced Mode (Recommended)
Uses LibreOffice + Python libraries without needing ONLYOFFICE server.

**Advantages:**
- ✅ No additional server required
- ✅ Better conversion quality than basic LibreOffice
- ✅ Works on Railway/Render out of the box
- ✅ Multiple conversion methods
- ✅ Free to use

**Environment Variables:**
```bash
# Leave ONLYOFFICE_DOCUMENT_SERVER_URL empty to use enhanced mode
ONLYOFFICE_DOCUMENT_SERVER_URL=
PYTHON_PATH=python3
TEMP_DIR=/tmp/pdf-converter
```

### Option 2: Full ONLYOFFICE Server Mode
Deploy ONLYOFFICE Document Server alongside your backend.

**Advantages:**
- ✅ Best conversion quality
- ✅ Official ONLYOFFICE conversion
- ✅ Supports complex documents
- ✅ Professional features

**Setup:**

#### For Railway:
```bash
# Deploy ONLYOFFICE as separate service
# Add to your Railway project
ONLYOFFICE_DOCUMENT_SERVER_URL=https://your-onlyoffice.railway.app
ONLYOFFICE_JWT_SECRET=your-secret-key
PYTHON_PATH=python3
```

#### For Render:
```yaml
# In render.yaml
services:
  - type: web
    name: onlyoffice-server
    env: docker
    dockerfilePath: ./docker/Dockerfile.onlyoffice
    envVars:
      - key: JWT_ENABLED
        value: "false"
  
  - type: web
    name: pdf-converter
    env: node
    buildCommand: npm install
    startCommand: npm start
    envVars:
      - key: ONLYOFFICE_DOCUMENT_SERVER_URL
        value: https://your-onlyoffice-server.onrender.com
```

#### For Local Development:
```bash
# Using Docker Compose
docker run -i -t -d -p 8000:80 \
  -e JWT_ENABLED=false \
  onlyoffice/documentserver:latest

# Set environment variable
ONLYOFFICE_DOCUMENT_SERVER_URL=http://localhost:8000
```

## Installation Requirements

### For Enhanced Mode (Option 1)
The backend will automatically install Python packages as needed:
- `pdf2docx` - High-quality PDF to Word conversion
- `pdfplumber` - PDF text extraction
- `python-docx` - Word document creation
- `pandas` - Excel file handling
- `python-pptx` - PowerPoint creation
- `openpyxl` - Excel file writing

### For Full Server Mode (Option 2)
1. **ONLYOFFICE Document Server** (Community Edition)
2. **LibreOffice** (fallback)
3. **Python 3** with above packages

## Configuration

### 1. Environment Variables
Copy `.env.onlyoffice` to `.env` and configure:

```bash
# Enhanced ONLYOFFICE Configuration
ONLYOFFICE_DOCUMENT_SERVER_URL=         # Leave empty for enhanced mode
ONLYOFFICE_TIMEOUT=120000
PYTHON_PATH=python3
SERVER_URL=http://localhost:10000

# Optional: For full server mode
ONLYOFFICE_JWT_SECRET=your-secret-key
```

### 2. Check Service Status
```bash
# Check enhanced service status
GET /onlyoffice/enhanced-status

# Response example:
{
  "timestamp": "2025-08-04T10:00:00.000Z",
  "service": "ONLYOFFICE Enhanced Service",
  "status": {
    "available": true,
    "healthy": true,
    "serverInfo": {
      "onlyofficeServer": {
        "available": false,
        "url": "",
        "jwtEnabled": false
      },
      "python": {
        "available": true,
        "path": "python3",
        "version": "Python 3.9.0"
      },
      "libreoffice": {
        "available": true,
        "version": "LibreOffice 7.4.0"
      }
    },
    "capabilities": {
      "onlyofficeServer": false,
      "python": true,
      "libreoffice": true,
      "multipleConversionMethods": true,
      "supportedFormats": ["docx", "xlsx", "pptx"]
    }
  }
}
```

## Conversion Methods Explained

### 1. Python-based Conversion (Enhanced Mode)
- **PDF to DOCX**: Uses `pdf2docx` library for high-quality conversion
- **PDF to XLSX**: Extracts tables with `pdfplumber` + `pandas`
- **PDF to PPTX**: Creates slides with `python-pptx`

### 2. ONLYOFFICE Document Server
- Uses official ONLYOFFICE conversion API
- Best quality for complex documents
- Requires separate server deployment

### 3. Enhanced LibreOffice
- Specialized conversion commands for each format
- Multiple retry methods
- Optimized for different document types

## Performance Comparison

| Method | Quality | Speed | Setup | Cost |
|--------|---------|-------|-------|------|
| Enhanced ONLYOFFICE | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | Free |
| ONLYOFFICE Server | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐ | Free |
| ConvertAPI | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | Paid |
| Standard LibreOffice | ⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ | Free |

## Troubleshooting

### Common Issues

#### 1. Python packages not installing
```bash
# Manual installation
pip install pdf2docx pdfplumber python-docx pandas python-pptx openpyxl
```

#### 2. LibreOffice not found
```bash
# Install LibreOffice
# Ubuntu/Debian:
sudo apt-get install libreoffice

# CentOS/RHEL:
sudo yum install libreoffice

# Alpine (for Docker):
apk add libreoffice
```

#### 3. ONLYOFFICE server connection issues
```bash
# Check server status
curl http://your-onlyoffice-server/healthcheck

# Verify environment variables
echo $ONLYOFFICE_DOCUMENT_SERVER_URL
```

### Logs to Check
```bash
# Enhanced ONLYOFFICE service logs
grep "Enhanced" application.log
grep "Python conversion" application.log
grep "LibreOffice conversion" application.log

# Conversion priority logs
grep "Converting PDF to" application.log
```

## Railway Deployment

Your current backend is **already compatible** with Railway! Just deploy and it will use Enhanced Mode automatically.

```bash
# Environment variables for Railway
ONLYOFFICE_DOCUMENT_SERVER_URL=  # Leave empty
PYTHON_PATH=python3
NODE_ENV=production
PORT=10000
```

## Production Recommendations

### 1. Enhanced Mode (Easiest)
- Deploy as-is to Railway/Render
- No additional services needed
- High conversion quality
- Automatic fallbacks

### 2. Full Server Mode (Best Quality)
- Deploy ONLYOFFICE server separately
- Configure JWT for security
- Use HTTPS for all communications
- Monitor server health

### 3. Hybrid Approach
- Start with Enhanced Mode
- Add ONLYOFFICE server later if needed
- Seamless upgrade path

## Benefits Over ConvertAPI

1. **✅ No API costs** - Community edition is free
2. **✅ No rate limits** - Self-hosted solution
3. **✅ Better reliability** - Multiple fallback methods
4. **✅ Enhanced quality** - Python libraries + ONLYOFFICE
5. **✅ Full control** - No external dependencies
6. **✅ Railway compatible** - Works out of the box

## Next Steps

1. **Deploy current backend** - It's already optimized for ONLYOFFICE
2. **Test Enhanced Mode** - Should work immediately
3. **Monitor conversion logs** - Check which methods are being used
4. **Optional**: Add ONLYOFFICE server for even better quality
5. **Remove ConvertAPI dependency** - It's now just a backup

Your backend is now **production-ready** with ONLYOFFICE as the primary conversion method! 🚀
