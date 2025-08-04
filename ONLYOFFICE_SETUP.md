# ONLYOFFICE Document Server Integration Guide

This guide explains how to integrate ONLYOFFICE Document Server (Community Edition) with your PDF converter backend for better PDF to Office format conversions.

## Overview

ONLYOFFICE Document Server provides superior PDF to Word/Excel/PowerPoint conversion compared to ConvertAPI, with the following benefits:

- **Free Community Edition** - No API costs or subscription fees
- **Better conversion quality** - Often produces higher quality results
- **No rate limits** - Unlike ConvertAPI's plan restrictions
- **Self-hosted** - Full control over your infrastructure
- **Railway/Render compatible** - Can be deployed alongside your backend

## Quick Setup for Development

### Option 1: Using Docker Compose (Recommended)

1. **Start ONLYOFFICE Document Server:**
   ```bash
   docker-compose -f docker-compose.dev.yml up -d
   ```

2. **Set environment variables:**
   ```bash
   export ONLYOFFICE_DOCUMENT_SERVER_URL=http://localhost:8000
   export ONLYOFFICE_TIMEOUT=120000
   ```

3. **Start your backend:**
   ```bash
   npm run build
   npm start
   ```

4. **Test the integration:**
   ```bash
   # Check ONLYOFFICE status
   curl http://localhost:10000/onlyoffice/status
   
   # The backend will now use ONLYOFFICE first, then fallback to ConvertAPI, then LibreOffice
   ```

### Option 2: Manual ONLYOFFICE Installation

1. **Install ONLYOFFICE Document Server:**
   ```bash
   # Ubuntu/Debian
   sudo apt-get install docker.io
   docker run -d -p 8000:80 --name onlyoffice-documentserver onlyoffice/documentserver
   
   # Or follow official installation guide:
   # https://github.com/ONLYOFFICE/DocumentServer
   ```

2. **Configure environment variables as above**

## Production Deployment

### For Railway

1. **Add ONLYOFFICE as a separate service:**
   - Create a new Railway service
   - Deploy ONLYOFFICE using Docker image: `onlyoffice/documentserver:latest`
   - Set environment variables:
     ```
     JWT_ENABLED=true
     JWT_SECRET=your-secure-jwt-secret
     ```

2. **Update your backend environment variables:**
   ```
   ONLYOFFICE_DOCUMENT_SERVER_URL=https://your-onlyoffice-service.railway.app
   ONLYOFFICE_JWT_SECRET=your-secure-jwt-secret
   ONLYOFFICE_TIMEOUT=120000
   ```

### For Render

1. **Deploy ONLYOFFICE as a separate Web Service:**
   - Create new Web Service in Render
   - Use Docker image: `onlyoffice/documentserver:latest`
   - Set environment variables:
     ```
     JWT_ENABLED=true
     JWT_SECRET=your-secure-jwt-secret
     ```

2. **Update your backend environment variables:**
   ```
   ONLYOFFICE_DOCUMENT_SERVER_URL=https://your-onlyoffice-service.onrender.com
   ONLYOFFICE_JWT_SECRET=your-secure-jwt-secret
   ONLYOFFICE_TIMEOUT=120000
   ```

## Environment Variables

| Variable | Description | Default | Required |
|----------|-------------|---------|----------|
| `ONLYOFFICE_DOCUMENT_SERVER_URL` | URL to ONLYOFFICE Document Server | - | Yes |
| `ONLYOFFICE_TIMEOUT` | Conversion timeout in milliseconds | 120000 | No |
| `ONLYOFFICE_JWT_SECRET` | JWT secret for secure communication | - | No (dev), Yes (prod) |

## Conversion Priority

The backend now uses this priority order for PDF conversions:

1. **ONLYOFFICE Document Server** (Primary)
2. **ConvertAPI** (Fallback if ONLYOFFICE fails)
3. **LibreOffice** (Final fallback)

## API Endpoints

### Check ONLYOFFICE Status
```bash
GET /onlyoffice/status
```

Response:
```json
{
  "timestamp": "2025-08-04T06:30:00.000Z",
  "onlyoffice": {
    "available": true,
    "healthy": true,
    "serverInfo": {
      "available": true,
      "healthy": true,
      "url": "http://localhost:8000",
      "jwtEnabled": false
    }
  }
}
```

## Troubleshooting

### ONLYOFFICE Not Available
- Check if `ONLYOFFICE_DOCUMENT_SERVER_URL` is set correctly
- Verify ONLYOFFICE server is running: `curl http://your-onlyoffice-url/healthcheck`
- Check logs: `docker logs onlyoffice-documentserver`

### Conversion Failures
- Check backend logs for detailed error messages
- Verify the PDF is not password-protected or corrupted
- Try different PDF files to isolate the issue

### Performance Issues
- Increase `ONLYOFFICE_TIMEOUT` for large files
- Monitor ONLYOFFICE server resources
- Consider using external database and Redis for production

## Resource Requirements

### Minimum (Development)
- CPU: 2 cores
- RAM: 4 GB
- Storage: 10 GB

### Recommended (Production)
- CPU: 4+ cores
- RAM: 8+ GB
- Storage: 50+ GB
- External database (MySQL/PostgreSQL)
- Redis for caching

## Security Considerations

### Development
- JWT can be disabled for local development
- Use HTTP for local testing

### Production
- **Always enable JWT** with a strong secret
- **Use HTTPS** for all communications
- **Restrict network access** to ONLYOFFICE server
- **Regular security updates**

## Advanced Configuration

### Custom Conversion Parameters
You can modify the ONLYOFFICE service to add custom conversion parameters:

```typescript
// In onlyoffice.service.ts
const conversionRequest = {
  async: false,
  filetype: 'pdf',
  key: this.generateConversionKey(filename),
  outputtype: targetFormat,
  title: filename,
  url: uploadUrl,
  // Custom parameters
  thumbnail: {
    aspect: 2,
    first: true,
    height: 100,
    width: 100
  }
};
```

### External Storage Integration
For production, consider integrating with cloud storage:

```typescript
// Example S3 integration
private async uploadFileToS3(buffer: Buffer, filename: string): Promise<string> {
  // Upload to S3 and return public URL
  // This URL will be accessible by ONLYOFFICE server
}
```

## Monitoring and Logging

Monitor conversion success rates and performance:

```bash
# Check conversion logs
tail -f logs/conversion.log | grep ONLYOFFICE

# Monitor resource usage
docker stats onlyoffice-documentserver
```

## Backup and Recovery

Regular backups of ONLYOFFICE data:

```bash
# Backup ONLYOFFICE data
docker run --rm -v onlyoffice_data:/data -v $(pwd):/backup alpine tar czf /backup/onlyoffice-backup.tar.gz /data

# Restore ONLYOFFICE data
docker run --rm -v onlyoffice_data:/data -v $(pwd):/backup alpine tar xzf /backup/onlyoffice-backup.tar.gz -C /
```
