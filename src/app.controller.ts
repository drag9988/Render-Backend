import { Controller, Post, Get, UploadedFile, UseInterceptors, Res, Body, Query, BadRequestException } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Throttle, SkipThrottle } from '@nestjs/throttler';
import { Response } from 'express';
import { AppService } from './app.service';
import { FileValidationService } from './file-validation.service';
import * as multer from 'multer';
import * as os from 'os';

@Controller()
export class AppController {
  constructor(
    private readonly appService: AppService,
    private readonly fileValidationService: FileValidationService
  ) {}

  // Health check endpoint
  @Get()
  @SkipThrottle() // No rate limiting for health checks
  healthCheck(@Res() res: Response) {
    try {
      const healthInfo = {
        status: 'ok',
        timestamp: new Date().toISOString(),
        uptime: process.uptime(),
        hostname: os.hostname(),
        network: Object.values(os.networkInterfaces())
          .flat()
          .filter(iface => iface && !iface.internal)
          .map(iface => ({
            address: iface.address,
            family: iface.family,
            netmask: iface.netmask,
          })),
        memory: {
          total: os.totalmem(),
          free: os.freemem(),
        },
        env: {
          nodeEnv: process.env.NODE_ENV || 'development',
          port: process.env.PORT || '3000',
        },
      };
      
      return res.status(200).json(healthInfo);
    } catch (error) {
      console.error('Health check error:', error.message);
      return res.status(500).json({ status: 'error', message: error.message });
    }
  }
  
  // Specific health check endpoint
  @Get('health')
  @SkipThrottle() // No rate limiting for health checks
  health(@Res() res: Response) {
    return res.status(200).json({ status: 'ok', timestamp: new Date().toISOString() });
  }

  // CORS test endpoint to verify CORS configuration
  @Get('cors-test')
  @SkipThrottle()
  corsTest(@Res() res: Response) {
    // Manually set CORS headers for testing
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Accept,Authorization,Content-Type,X-Requested-With,Range,Origin');
    
    return res.status(200).json({ 
      message: 'CORS test successful',
      timestamp: new Date().toISOString(),
      headers: {
        origin: res.req.get('Origin') || 'none',
        userAgent: res.req.get('User-Agent') || 'none',
        method: res.req.method
      }
    });
  }

  @Post('cors-test')
  @SkipThrottle()
  corsTestPost(@Res() res: Response, @Body() body: any) {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Accept,Authorization,Content-Type,X-Requested-With,Range,Origin');
    
    return res.status(200).json({ 
      message: 'CORS POST test successful',
      receivedBody: body,
      timestamp: new Date().toISOString()
    });
  }

  // Word to PDF conversion
  @Post('convert-word-to-pdf')
  @Throttle({ default: { limit: 20, ttl: 86400000 } }) // 20 requests per day for office to PDF
  @UseInterceptors(FileInterceptor('file'))
  async convertWordToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }

      // Validate Word file
      const validation = this.fileValidationService.validateFile(file, 'word');
      if (!validation.isValid) {
        return res.status(400).json({ 
          error: 'File validation failed', 
          details: validation.errors 
        });
      }

      // Update file with sanitized filename
      file.originalname = validation.sanitizedFilename;
      
      // Use LibreOffice directly for Word to PDF conversion
      const output = await this.appService.convertOfficeToPdf(file);
      res.set({ 'Content-Type': 'application/pdf' });
      res.send(output);
    } catch (error) {
      console.error('Word to PDF conversion error:', error.message);
      if (error instanceof BadRequestException) {
        return res.status(400).json({ error: 'Invalid file', message: error.message });
      }
      res.status(500).json({ error: 'Failed to convert Word to PDF', message: error.message });
    }
  }

  // Excel to PDF conversion
  @Post('convert-excel-to-pdf')
  @Throttle({ default: { limit: 20, ttl: 86400000 } }) // 20 requests per day for office to PDF
  @UseInterceptors(FileInterceptor('file'))
  async convertExcelToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }

      // Validate Excel file
      const validation = this.fileValidationService.validateFile(file, 'excel');
      if (!validation.isValid) {
        return res.status(400).json({ 
          error: 'File validation failed', 
          details: validation.errors 
        });
      }

      // Update file with sanitized filename
      file.originalname = validation.sanitizedFilename;
      
      // Use LibreOffice directly for Excel to PDF conversion
      const output = await this.appService.convertOfficeToPdf(file);
      res.set({ 'Content-Type': 'application/pdf' });
      res.send(output);
    } catch (error) {
      console.error('Excel to PDF conversion error:', error.message);
      if (error instanceof BadRequestException) {
        return res.status(400).json({ error: 'Invalid file', message: error.message });
      }
      res.status(500).json({ error: 'Failed to convert Excel to PDF', message: error.message });
    }
  }

  // PowerPoint to PDF conversion
  @Post('convert-ppt-to-pdf')
  @Throttle({ default: { limit: 20, ttl: 86400000 } }) // 20 requests per day for office to PDF
  @UseInterceptors(FileInterceptor('file'))
  async convertPptToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }

      // Validate PowerPoint file
      const validation = this.fileValidationService.validateFile(file, 'powerpoint');
      if (!validation.isValid) {
        return res.status(400).json({ 
          error: 'File validation failed', 
          details: validation.errors 
        });
      }

      // Update file with sanitized filename
      file.originalname = validation.sanitizedFilename;
      
      // Use LibreOffice directly for PowerPoint to PDF conversion
      const output = await this.appService.convertOfficeToPdf(file);
      res.set({ 'Content-Type': 'application/pdf' });
      res.send(output);
    } catch (error) {
      console.error('PowerPoint to PDF conversion error:', error.message);
      if (error instanceof BadRequestException) {
        return res.status(400).json({ error: 'Invalid file', message: error.message });
      }
      res.status(500).json({ error: 'Failed to convert PowerPoint to PDF', message: error.message });
    }
  }

  // PDF to Word conversion
  @Post('convert-pdf-to-word')
  @Throttle({ default: { limit: 5, ttl: 86400000 } }) // 5 requests per day
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToWord(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }

      // Validate PDF file
      const validation = this.fileValidationService.validateFile(file, 'pdf');
      if (!validation.isValid) {
        return res.status(400).json({ 
          error: 'PDF validation failed', 
          details: validation.errors 
        });
      }

      // Update file with sanitized filename
      file.originalname = validation.sanitizedFilename;
      
      console.log(`Converting PDF to Word: ${file.originalname}, size: ${file.size} bytes`);
      // Use ONLYOFFICE Enhanced Service for PDF to Word conversion
      const output = await this.appService.convertPdfToOffice(file, 'docx');
      
      res.set({ 
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '.docx')}"` 
      });
      res.send(output);
    } catch (error) {
      console.error('PDF to Word conversion error:', error.message);
      
      // Provide more specific error messages
      if (error.message.includes('timeout')) {
        return res.status(408).json({ 
          error: 'Conversion timeout', 
          message: 'The PDF conversion is taking too long. Please try with a smaller or simpler PDF.' 
        });
      } else if (error.message.includes('complex formatting')) {
        return res.status(422).json({ 
          error: 'Complex PDF format', 
          message: 'This PDF contains complex formatting that cannot be converted. Try using a text-based PDF instead of a scanned document.',
          suggestions: [
            'Ensure the PDF contains selectable text (not a scanned image)',
            'Try with a simpler PDF with basic formatting',
            'Consider using a PDF with fewer images and complex layouts'
          ]
        });
      } else if (error.message.includes('scanned PDF')) {
        return res.status(422).json({ 
          error: 'Scanned PDF detected', 
          message: 'This appears to be a scanned PDF (image-based) which cannot be converted to editable Word format.',
          suggestions: [
            'Use OCR software to convert the scanned PDF to text first',
            'Try with a text-based PDF instead',
            'Export the original document directly to Word if possible'
          ]
        });
      } else if (error.message.includes('insufficient text')) {
        return res.status(422).json({ 
          error: 'Insufficient text content', 
          message: 'This PDF does not contain enough text content for conversion.',
          suggestions: [
            'Verify the PDF contains readable text',
            'Check if the PDF is password-protected',
            'Try with a different PDF file'
          ]
        });
      } else {
        return res.status(500).json({ 
          error: 'Failed to convert PDF to Word', 
          message: error.message 
        });
      }
    }
  }

  // PDF to Excel conversion
  @Post('convert-pdf-to-excel')
  @Throttle({ default: { limit: 5, ttl: 86400000 } }) // 5 requests per day
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToExcel(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }

      // Validate PDF file
      const validation = this.fileValidationService.validateFile(file, 'pdf');
      if (!validation.isValid) {
        return res.status(400).json({ 
          error: 'PDF validation failed', 
          details: validation.errors 
        });
      }

      // Update file with sanitized filename
      file.originalname = validation.sanitizedFilename;
      
      console.log(`🚀 Starting ULTIMATE PDF to Excel conversion: ${file.originalname}, size: ${file.size} bytes`);
      console.log(`📊 Using Enhanced ONLYOFFICE Service with advanced table detection...`);
      
      // Use ONLYOFFICE Enhanced Service for PDF to Excel conversion
      const output = await this.appService.convertPdfToOffice(file, 'xlsx');
      
      console.log(`✅ PDF to Excel conversion successful: ${output.length} bytes output`);
      
      res.set({ 
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '.xlsx')}"`,
        'X-Conversion-Method': 'Enhanced-ONLYOFFICE-Service',
        'X-Output-Size': output.length.toString()
      });
      res.send(output);
    } catch (error) {
      console.error('❌ PDF to Excel conversion error:', error.message);
      
      if (error.message.includes('timeout')) {
        return res.status(408).json({ 
          error: 'Conversion timeout', 
          message: 'The PDF to Excel conversion is taking too long. This usually happens with large or complex PDFs.',
          suggestions: [
            'Try with a smaller PDF file (under 10MB)',
            'Split large PDFs into smaller sections',
            'Ensure the PDF contains clear tabular data for better conversion'
          ],
          recommendations: {
            'file_size': 'Keep PDF files under 10MB for optimal conversion speed',
            'content_type': 'PDFs with clear tables and structured data convert best',
            'retry': 'Please try again in a few minutes'
          }
        });
      } else if (error.message.includes('complex formatting') || error.message.includes('complex PDF')) {
        return res.status(422).json({ 
          error: 'Complex PDF format detected', 
          message: 'This PDF has complex formatting that makes table extraction challenging.',
          details: 'The Enhanced ONLYOFFICE Service attempted multiple conversion methods but could not extract reliable tabular data.',
          suggestions: [
            'Ensure the PDF contains clear tables with visible borders or consistent spacing',
            'Try with a PDF that has simple table layouts',
            'Consider using a PDF with text-based tables rather than image-based ones',
            'If the PDF has multiple tables, they should be well-separated'
          ],
          alternatives: [
            'Export the data directly from the original source (e.g., Excel, database)',
            'Use specialized PDF table extraction software',
            'Convert the PDF to text first, then organize the data manually'
          ]
        });
      } else if (error.message.includes('scanned PDF') || error.message.includes('image')) {
        return res.status(422).json({ 
          error: 'Scanned PDF or image-based content detected', 
          message: 'This appears to be a scanned PDF or contains primarily images, which cannot be directly converted to Excel.',
          details: 'The system detected that this PDF likely contains scanned images rather than selectable text and tables.',
          suggestions: [
            'Use OCR (Optical Character Recognition) software first to make the PDF text-selectable',
            'Try with a text-based PDF that allows text selection',
            'Export data directly from the original source if possible',
            'Use specialized OCR tools that can recognize table structures'
          ],
          tools_recommended: [
            'Adobe Acrobat Pro (OCR functionality)',
            'Tabula (for table extraction from PDFs)',
            'ABBYY FineReader (advanced OCR)',
            'Online OCR services with table recognition'
          ]
        });
      } else if (error.message.includes('no tables') || error.message.includes('no data')) {
        return res.status(422).json({ 
          error: 'No tabular data found', 
          message: 'The PDF was processed successfully, but no table structures could be detected.',
          details: 'The Enhanced ONLYOFFICE Service analyzed the PDF but could not identify organized data suitable for Excel format.',
          suggestions: [
            'Verify that the PDF contains tables or structured data',
            'Check if the data is organized in columns and rows',
            'Ensure tables have clear boundaries or consistent spacing',
            'Try with a PDF that has more traditional table layouts'
          ],
          conversion_attempted: [
            'Advanced table detection algorithms',
            'Text-based data extraction',
            'Multiple conversion methods including Tabula, Camelot, and pdfplumber',
            'Intelligent data structure analysis'
          ]
        });
      } else if (error.message.includes('All premium conversion methods failed')) {
        return res.status(500).json({
          error: 'Advanced conversion methods exhausted',
          message: 'All available high-quality conversion methods have been attempted.',
          details: 'The Enhanced ONLYOFFICE Service tried multiple premium conversion approaches including ONLYOFFICE Document Server, advanced Python libraries, and enhanced LibreOffice methods.',
          methods_attempted: [
            'ONLYOFFICE Document Server (if configured)',
            'Tabula (advanced table detection)',
            'Camelot (premium table extraction)', 
            'pdfplumber (intelligent text extraction)',
            'Enhanced LibreOffice with Calc optimization',
            'PyMuPDF with smart data structuring'
          ],
          recommendations: [
            'Deploy ONLYOFFICE Document Server for best results',
            'Try with a simpler PDF containing clear table structures',
            'Consider manual data extraction for complex documents',
            'Use the original application to export data directly'
          ],
          technical_support: 'Contact technical support if this is a recurring issue'
        });
      } else if (error.message.includes('File validation failed')) {
        return res.status(400).json({
          error: 'PDF file validation failed',
          message: 'The uploaded file failed security and format validation.',
          details: error.message,
          requirements: [
            'File must be a valid PDF format',
            'File size must be under 50MB',
            'File must not be password-protected',
            'File must contain readable content'
          ]
        });
      } else {
        return res.status(500).json({ 
          error: 'PDF to Excel conversion failed', 
          message: 'An unexpected error occurred during the conversion process.',
          details: error.message,
          support_info: {
            'conversion_service': 'Enhanced ONLYOFFICE Service',
            'error_type': 'Internal processing error',
            'recommended_action': 'Please try again or contact support if the issue persists'
          }
        });
      }
    }
  }

  // PDF to PowerPoint conversion
  @Post('convert-pdf-to-ppt')
  @Throttle({ default: { limit: 5, ttl: 86400000 } }) // 5 requests per day
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToPpt(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }

      // Validate PDF file
      const validation = this.fileValidationService.validateFile(file, 'pdf');
      if (!validation.isValid) {
        return res.status(400).json({ 
          error: 'PDF validation failed', 
          details: validation.errors 
        });
      }

      // Update file with sanitized filename
      file.originalname = validation.sanitizedFilename;
      
      console.log(`Converting PDF to PowerPoint: ${file.originalname}, size: ${file.size} bytes`);
      // Use ONLYOFFICE Enhanced Service for PDF to PowerPoint conversion
      const output = await this.appService.convertPdfToOffice(file, 'pptx');
      
      res.set({ 
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '.pptx')}"` 
      });
      res.send(output);
    } catch (error) {
      console.error('PDF to PowerPoint conversion error:', error.message);
      
      if (error.message.includes('timeout')) {
        return res.status(408).json({ 
          error: 'Conversion timeout', 
          message: 'The PDF conversion is taking too long. Please try with a smaller or simpler PDF.' 
        });
      } else if (error.message.includes('complex formatting')) {
        return res.status(422).json({ 
          error: 'Complex PDF format', 
          message: 'This PDF contains complex formatting that may not convert well to PowerPoint slides.',
          suggestions: [
            'Try with a PDF that has clear page-by-page content',
            'Ensure the PDF has simple layout without complex graphics',
            'Consider breaking down the PDF into smaller sections'
          ]
        });
      } else if (error.message.includes('scanned PDF')) {
        return res.status(422).json({ 
          error: 'Scanned PDF detected', 
          message: 'This appears to be a scanned PDF which cannot be converted to PowerPoint format.',
          suggestions: [
            'Use OCR software to make the PDF text-selectable first',
            'Try with a text-based PDF',
            'Export slides directly from the original presentation if possible'
          ]
        });
      } else {
        return res.status(500).json({ 
          error: 'Failed to convert PDF to PowerPoint', 
          message: error.message 
        });
      }
    }
  }

  // PDF analysis endpoint - check if PDF is suitable for conversion
  @Post('analyze-pdf')
  @Throttle({ default: { limit: 10, ttl: 86400000 } }) // 10 requests per day for analysis
  @UseInterceptors(FileInterceptor('file'))
  async analyzePdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      // Validate file type
      if (file.mimetype !== 'application/pdf') {
        return res.status(400).json({ error: 'Uploaded file is not a PDF' });
      }

      // Save file temporarily for analysis
      const timestamp = Date.now();
      const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
      
      // Ensure temp directory exists
      try {
        await require('fs').promises.mkdir(tempDir, { recursive: true });
        await require('fs').promises.chmod(tempDir, 0o777);
      } catch (dirError) {
        console.error(`Failed to create temp directory ${tempDir}: ${dirError.message}`);
        return res.status(500).json({ error: 'Server configuration error', message: 'Cannot access temporary storage' });
      }
      
      const tempInput = `${tempDir}/${timestamp}_analysis.pdf`;
      
      await require('fs').promises.writeFile(tempInput, file.buffer);
      
      // Analyze the PDF
      const analysis = await this.appService.analyzePdfFile(tempInput);
      
      // Clean up
      await require('fs').promises.unlink(tempInput).catch(() => {});
      
      res.json({
        filename: file.originalname,
        size: file.size,
        analysis: analysis,
        recommendations: this.getConversionRecommendations(analysis)
      });
      
    } catch (error) {
      console.error('PDF analysis error:', error.message);
      res.status(500).json({ error: 'Failed to analyze PDF', message: error.message });
    }
  }

  private getConversionRecommendations(analysis: any): any {
    const recommendations = {
      canConvertToWord: true,
      canConvertToExcel: true,
      canConvertToPowerPoint: true,
      warnings: [],
      suggestions: []
    };

    if (analysis.isScanned) {
      recommendations.canConvertToWord = false;
      recommendations.canConvertToExcel = false;
      recommendations.canConvertToPowerPoint = false;
      recommendations.warnings.push('This appears to be a scanned PDF (image-based)');
      recommendations.suggestions.push('Use OCR software to make the PDF text-selectable first');
    }

    if (analysis.isProtected) {
      recommendations.canConvertToWord = false;
      recommendations.canConvertToExcel = false;
      recommendations.canConvertToPowerPoint = false;
      recommendations.warnings.push('This PDF appears to be password-protected or restricted');
      recommendations.suggestions.push('Remove password protection before conversion');
    }

    if (analysis.hasComplexLayout) {
      recommendations.warnings.push('Complex layout detected - conversion quality may vary');
      recommendations.suggestions.push('Simpler PDFs generally convert better');
    }

    if (analysis.pageCount > 50) {
      recommendations.warnings.push('Large PDF detected - conversion may take longer');
      recommendations.suggestions.push('Consider splitting into smaller files for faster processing');
    }

    return recommendations;
  }

  // PDF compression with quality options
  @Post('compress-pdf')
  @Throttle({ default: { limit: 15, ttl: 86400000 } }) // 15 requests per day for compression
  @UseInterceptors(FileInterceptor('file', {
    limits: {
      fileSize: 50 * 1024 * 1024, // 50MB file size limit
    },
  }))
  async compressPdf(
    @UploadedFile() file: Express.Multer.File,
    @Query('quality') quality: string = 'moderate',
    @Res() res: Response,
  ) {
    // CRITICAL DEBUG: Log as soon as request hits the endpoint
    console.log(`🔥 COMPRESS-PDF REQUEST RECEIVED at ${new Date().toISOString()}`);
    console.log(`🔥 Request headers:`, JSON.stringify(Object.keys(res.req.headers || {})));
    console.log(`🔥 File received:`, file ? `YES (${file.size} bytes)` : 'NO');
    console.log(`🔥 Quality parameter:`, quality);
    
    try {
      if (!file) {
        console.log(`🔥 ERROR: No file uploaded`);
        return res.status(400).json({ error: 'No PDF file uploaded' });
      }

      // Additional file size check
      if (file.size > 50 * 1024 * 1024) {
        return res.status(400).json({ 
          error: 'File too large', 
          message: 'PDF file size must be less than 50MB',
          currentSize: `${(file.size / (1024 * 1024)).toFixed(2)}MB`,
          maxSize: '50MB'
        });
      }

      // Validate PDF file
      const validation = this.fileValidationService.validateFile(file, 'pdf');
      if (!validation.isValid) {
        return res.status(400).json({ 
          error: 'PDF validation failed', 
          details: validation.errors 
        });
      }

      // Update file with sanitized filename
      file.originalname = validation.sanitizedFilename;
      
      console.log(`Starting PDF compression: ${file.originalname}, size: ${file.size} bytes, quality: ${quality}`);
      
      // Enhanced logging for image-heavy PDFs
      const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
      if (file.size > 10 * 1024 * 1024) { // >10MB
        console.log(`Large PDF detected (${sizeMB}MB) - likely contains high-resolution images from mobile camera`);
        console.log(`Using extended timeout for image-heavy PDF compression...`);
      }
      
      // Log the quality setting being used
      console.log(`Compression quality setting: ${quality}`);
      
      const startTime = Date.now();
      const output = await this.appService.compressPdf(file, quality);
      const processingTime = (Date.now() - startTime) / 1000;
      
      const outputSizeMB = (output.length / (1024 * 1024)).toFixed(2);
      const compressionRatio = ((1 - output.length / file.size) * 100).toFixed(1);
      
      console.log(`PDF compression completed: ${file.originalname}`);
      console.log(`  Original: ${sizeMB}MB`);
      console.log(`  Compressed: ${outputSizeMB}MB`);
      console.log(`  Compression: ${compressionRatio}%`);
      console.log(`  Processing time: ${processingTime}s`);
      
      res.set({ 'Content-Type': 'application/pdf' });
      res.send(output);
    } catch (error) {
      console.error('PDF compression error:', error.message);
      console.error('Error stack:', error.stack);
      
      // Check for multer file size errors first
      if (error.code === 'LIMIT_FILE_SIZE') {
        return res.status(400).json({ 
          error: 'File too large', 
          message: 'PDF file size exceeds the 50MB limit',
          maxSize: '50MB'
        });
      }
      
      if (error instanceof BadRequestException) {
        return res.status(400).json({ error: 'Invalid file', message: error.message });
      }
      
      // Enhanced error handling for image-heavy PDFs and timeouts
      if (error.message.includes('timeout') || error.message.includes('timed out')) {
        const isImageHeavyError = error.message.includes('mobile camera') || error.message.includes('high-resolution');
        
        return res.status(408).json({ 
          error: 'Compression timeout', 
          message: isImageHeavyError 
            ? 'This PDF contains high-resolution images (likely mobile camera photos) that take too long to compress.'
            : 'The PDF compression is taking too long. This often happens with large or complex PDFs.',
          suggestions: [
            'Try using "low" quality setting for faster compression',
            'Reduce image resolution before creating the PDF',
            'Try with a smaller PDF file (under 10MB for best results)',
            'Consider compressing images separately before adding to PDF',
            'Split large PDFs into smaller files'
          ]
        });
      } else if (error.message.includes('image-heavy') || error.message.includes('high-resolution') || error.message.includes('mobile camera')) {
        return res.status(422).json({ 
          error: 'High-resolution images detected', 
          message: 'This PDF contains high-resolution images (likely from mobile camera photos) that are difficult to compress.',
          suggestions: [
            'Use "low" quality setting for better compression of mobile photos',
            'Reduce image quality/resolution before creating the PDF',
            'Try compressing images to JPEG format before adding to PDF',
            'Consider using image editing software to reduce file sizes first'
          ]
        });
      } else if (error.message.includes('service is not available') || error.message.includes('command not found') || error.message.includes('ghostscript')) {
        return res.status(503).json({ 
          error: 'Service unavailable', 
          message: 'PDF compression service is temporarily unavailable. The Ghostscript service may be down.' 
        });
      } else if (error.message.includes('All compression methods failed')) {
        return res.status(422).json({ 
          error: 'Compression failed', 
          message: 'This PDF could not be compressed using any available method. It may contain very complex content or be corrupted.',
          suggestions: [
            'Try with a different PDF file',
            'Check if the PDF is password-protected',
            'Try converting the PDF to images and back to PDF first',
            'Ensure the PDF is not corrupted',
            'Try with a smaller, simpler PDF'
          ]
        });
      } else if (error.message.includes('poppler-utils') || error.message.includes('pdfinfo') || error.message.includes('pdfimages')) {
        // Handle analysis errors gracefully
        return res.status(500).json({ 
          error: 'PDF analysis failed', 
          message: 'Could not analyze PDF content, but compression may still work.',
          details: 'PDF content analysis tools are not available',
          suggestions: [
            'Try compression anyway - it may still work',
            'Use "low" quality setting',
            'Try with a smaller PDF file'
          ]
        });
      }
      
      res.status(500).json({ 
        error: 'Failed to compress PDF', 
        message: 'An unexpected error occurred during compression.',
        details: error.message,
        suggestions: [
          'Try with a different PDF file',
          'Use "low" quality setting for mobile camera photos',
          'Ensure the PDF is not corrupted or password-protected',
          'Try with a smaller file (under 50MB)'
        ]
      });
    }
  }

  // ConvertAPI status endpoint
  @Get('convertapi/status')
  @SkipThrottle() // No rate limiting for status checks
  async getConvertApiStatus(@Res() res: Response) {
    try {
      const status = await this.appService.getConvertApiStatus();
      return res.status(200).json({
        timestamp: new Date().toISOString(),
        convertapi: status
      });
    } catch (error) {
      console.error('ConvertAPI status check error:', error.message);
      return res.status(500).json({ 
        error: 'Failed to check ConvertAPI status', 
        message: error.message 
      });
    }
  }

  // ONLYOFFICE status endpoint
  @Get('onlyoffice/status')
  @SkipThrottle() // No rate limiting for status checks
  async getOnlyOfficeStatus(@Res() res: Response) {
    try {
      const status = await this.appService.getOnlyOfficeStatus();
      return res.status(200).json({
        timestamp: new Date().toISOString(),
        onlyoffice: status
      });
    } catch (error) {
      console.error('ONLYOFFICE status check error:', error.message);
      return res.status(500).json({ 
        error: 'Failed to check ONLYOFFICE status', 
        message: error.message 
      });
    }
  }

  // Enhanced ONLYOFFICE status endpoint with detailed info
  @Get('onlyoffice/enhanced-status')
  @SkipThrottle() // No rate limiting for status checks
  async getEnhancedOnlyOfficeStatus(@Res() res: Response) {
    try {
      const status = await this.appService.getEnhancedOnlyOfficeStatus();
      return res.status(200).json({
        timestamp: new Date().toISOString(),
        service: 'ONLYOFFICE Enhanced Service',
        status: status
      });
    } catch (error) {
      console.error('Enhanced ONLYOFFICE status check error:', error.message);
      return res.status(500).json({ 
        error: 'Failed to check Enhanced ONLYOFFICE status', 
        message: error.message 
      });
    }
  }

  // PDF password protection
  @Post('add-password-to-pdf')
  @Throttle({ default: { limit: 10, ttl: 86400000 } }) // 10 requests per day for password protection
  @UseInterceptors(FileInterceptor('file'))
  async addPasswordToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Body('password') password: string,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No PDF file uploaded' });
      }

      if (!password) {
        return res.status(400).json({ error: 'Password is required' });
      }

      // Validate PDF file
      const validation = this.fileValidationService.validateFile(file, 'pdf');
      if (!validation.isValid) {
        return res.status(400).json({ 
          error: 'PDF validation failed', 
          details: validation.errors 
        });
      }

      // Update file with sanitized filename
      file.originalname = validation.sanitizedFilename;
      
      console.log(`Adding password protection to PDF: ${file.originalname}, size: ${file.size} bytes`);
      const output = await this.appService.addPasswordToPdf(file, password);
      
      res.set({ 
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '_protected.pdf')}"` 
      });
      res.send(output);
    } catch (error) {
      console.error('PDF password protection error:', error.message);
      
      if (error instanceof BadRequestException) {
        return res.status(400).json({ error: 'Invalid input', message: error.message });
      }
      
      if (error.message.includes('timeout')) {
        return res.status(408).json({ 
          error: 'Password protection timeout', 
          message: 'The PDF password protection is taking too long. Please try with a smaller PDF.' 
        });
      } else if (error.message.includes('LibreOffice')) {
        return res.status(503).json({ 
          error: 'Service unavailable', 
          message: 'LibreOffice service is not available. Please try again later.' 
        });
      }
      
      res.status(500).json({ 
        error: 'Failed to add password protection to PDF', 
        message: error.message 
      });
    }
  }
}