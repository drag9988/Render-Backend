import { Controller, Post, Get, UploadedFile, UseInterceptors, Res, Body, Query, Req, Logger } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Throttle, SkipThrottle } from '@nestjs/throttler';
import { Response, Request } from 'express';
import { AppService } from './app.service';
import { SecurityService } from './security.service';
import * as multer from 'multer';
import * as os from 'os';

@Controller()
export class AppController {
  private readonly logger = new Logger(AppController.name);

  constructor(
    private readonly appService: AppService,
    private readonly securityService: SecurityService
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

  // Word to PDF conversion
  @Post('convert-word-to-pdf')
  @Throttle({ default: { limit: 20, ttl: 86400000 } }) // 20 requests per day for office to PDF
  @UseInterceptors(FileInterceptor('file'))
  async convertWordToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Req() req: Request,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      // Security validation
      this.validateFileAndRequest(file, req, 'office');
      
      const sanitizedFilename = this.securityService.sanitizeFilename(file.originalname);
      this.logger.log(`Converting Word to PDF: ${sanitizedFilename}, size: ${file.size} bytes`);
      
      const output = await this.appService.convertLibreOffice(file, 'pdf');
      res.set({ 
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${sanitizedFilename.replace(/\.(docx?|txt|rtf)$/i, '.pdf')}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY'
      });
      res.send(output);
    } catch (error) {
      this.logger.error('Word to PDF conversion error:', error.message);
      
      if (error.message.includes('Security validation failed')) {
        return res.status(400).json({ 
          error: 'Security validation failed', 
          message: error.message 
        });
      } else {
        res.status(500).json({ error: 'Failed to convert Word to PDF', message: error.message });
      }
    }
  }

  // Excel to PDF conversion
  @Post('convert-excel-to-pdf')
  @Throttle({ default: { limit: 20, ttl: 86400000 } }) // 20 requests per day for office to PDF
  @UseInterceptors(FileInterceptor('file'))
  async convertExcelToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Req() req: Request,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      // Security validation
      this.validateFileAndRequest(file, req, 'office');
      
      const sanitizedFilename = this.securityService.sanitizeFilename(file.originalname);
      this.logger.log(`Converting Excel to PDF: ${sanitizedFilename}, size: ${file.size} bytes`);
      
      const output = await this.appService.convertLibreOffice(file, 'pdf');
      res.set({ 
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${sanitizedFilename.replace(/\.(xlsx?|csv)$/i, '.pdf')}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY'
      });
      res.send(output);
    } catch (error) {
      this.logger.error('Excel to PDF conversion error:', error.message);
      
      if (error.message.includes('Security validation failed')) {
        return res.status(400).json({ 
          error: 'Security validation failed', 
          message: error.message 
        });
      } else {
        res.status(500).json({ error: 'Failed to convert Excel to PDF', message: error.message });
      }
    }
  }

  // PowerPoint to PDF conversion
  @Post('convert-ppt-to-pdf')
  @Throttle({ default: { limit: 20, ttl: 86400000 } }) // 20 requests per day for office to PDF
  @UseInterceptors(FileInterceptor('file'))
  async convertPptToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Req() req: Request,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      // Security validation
      this.validateFileAndRequest(file, req, 'office');
      
      const sanitizedFilename = this.securityService.sanitizeFilename(file.originalname);
      this.logger.log(`Converting PowerPoint to PDF: ${sanitizedFilename}, size: ${file.size} bytes`);
      
      const output = await this.appService.convertLibreOffice(file, 'pdf');
      res.set({ 
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${sanitizedFilename.replace(/\.(pptx?|odp)$/i, '.pdf')}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY'
      });
      res.send(output);
    } catch (error) {
      console.error('PowerPoint to PDF conversion error:', error.message);
      res.status(500).json({ error: 'Failed to convert PowerPoint to PDF', message: error.message });
    }
  }

  // PDF to Word conversion
  @Post('convert-pdf-to-word')
  @Throttle({ default: { limit: 5, ttl: 86400000 } }) // 5 requests per day
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToWord(
    @UploadedFile() file: Express.Multer.File,
    @Req() req: Request,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      // Security validation
      this.validateFileAndRequest(file, req, 'pdf');
      
      // Additional PDF-specific validations
      if (file.mimetype !== 'application/pdf') {
        return res.status(400).json({ 
          error: 'Invalid file type', 
          message: 'Only PDF files are allowed for this conversion',
          receivedType: file.mimetype 
        });
      }

      // Validate file size (PDFs larger than 50MB may cause issues)
      if (file.size > 50 * 1024 * 1024) {
        return res.status(400).json({ 
          error: 'File too large', 
          message: 'PDF file must be smaller than 50MB',
          fileSize: `${Math.round(file.size / 1024 / 1024)}MB` 
        });
      }
      
      // Sanitize filename for logging
      const sanitizedFilename = this.securityService.sanitizeFilename(file.originalname);
      this.logger.log(`Converting PDF to Word: ${sanitizedFilename}, size: ${file.size} bytes`);
      
      const output = await this.appService.convertLibreOffice(file, 'docx');
      
      res.set({ 
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${sanitizedFilename.replace('.pdf', '.docx')}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY'
      });
      res.send(output);
    } catch (error) {
      this.logger.error('PDF to Word conversion error:', error.message);
      
      // Provide more specific error messages
      if (error.message.includes('Security validation failed')) {
        return res.status(400).json({ 
          error: 'Security validation failed', 
          message: error.message 
        });
      } else if (error.message.includes('timeout')) {
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
    @Req() req: Request,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      // Security validation
      this.validateFileAndRequest(file, req, 'pdf');
      
      // Validate file type
      if (file.mimetype !== 'application/pdf') {
        return res.status(400).json({ 
          error: 'Invalid file type', 
          message: 'Only PDF files are allowed for this conversion',
          receivedType: file.mimetype 
        });
      }

      // Validate file size
      if (file.size > 50 * 1024 * 1024) {
        return res.status(400).json({ 
          error: 'File too large', 
          message: 'PDF file must be smaller than 50MB',
          fileSize: `${Math.round(file.size / 1024 / 1024)}MB` 
        });
      }
      
      const sanitizedFilename = this.securityService.sanitizeFilename(file.originalname);
      this.logger.log(`Converting PDF to Excel: ${sanitizedFilename}, size: ${file.size} bytes`);
      
      const output = await this.appService.convertLibreOffice(file, 'xlsx');
      
      res.set({ 
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${sanitizedFilename.replace('.pdf', '.xlsx')}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY'
      });
      res.send(output);
    } catch (error) {
      this.logger.error('PDF to Excel conversion error:', error.message);
      
      if (error.message.includes('Security validation failed')) {
        return res.status(400).json({ 
          error: 'Security validation failed', 
          message: error.message 
        });
      } else if (error.message.includes('timeout')) {
        return res.status(408).json({ 
          error: 'Conversion timeout', 
          message: 'The PDF conversion is taking too long. Please try with a smaller or simpler PDF.' 
        });
      } else if (error.message.includes('complex formatting')) {
        return res.status(422).json({ 
          error: 'Complex PDF format', 
          message: 'This PDF may not contain tabular data suitable for Excel conversion.',
          suggestions: [
            'Ensure the PDF contains tables or structured data',
            'Try with a PDF that has clear tabular layout',
            'Consider manual copy-paste for complex data structures'
          ]
        });
      } else if (error.message.includes('scanned PDF')) {
        return res.status(422).json({ 
          error: 'Scanned PDF detected', 
          message: 'This appears to be a scanned PDF which cannot be converted to Excel format.',
          suggestions: [
            'Use OCR software first to make the PDF text-selectable',
            'Try with a text-based PDF containing tables',
            'Export data directly from the original source if possible'
          ]
        });
      } else {
        return res.status(500).json({ 
          error: 'Failed to convert PDF to Excel', 
          message: error.message 
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
    @Req() req: Request,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      // Security validation
      this.validateFileAndRequest(file, req, 'pdf');
      
      // Validate file type
      if (file.mimetype !== 'application/pdf') {
        return res.status(400).json({ 
          error: 'Invalid file type', 
          message: 'Only PDF files are allowed for this conversion',
          receivedType: file.mimetype 
        });
      }

      // Validate file size
      if (file.size > 50 * 1024 * 1024) {
        return res.status(400).json({ 
          error: 'File too large', 
          message: 'PDF file must be smaller than 50MB',
          fileSize: `${Math.round(file.size / 1024 / 1024)}MB` 
        });
      }
      
      const sanitizedFilename = this.securityService.sanitizeFilename(file.originalname);
      this.logger.log(`Converting PDF to PowerPoint: ${sanitizedFilename}, size: ${file.size} bytes`);
      
      const output = await this.appService.convertLibreOffice(file, 'pptx');
      
      res.set({ 
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${sanitizedFilename.replace('.pdf', '.pptx')}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY'
      });
      res.send(output);
    } catch (error) {
      this.logger.error('PDF to PowerPoint conversion error:', error.message);
      
      if (error.message.includes('Security validation failed')) {
        return res.status(400).json({ 
          error: 'Security validation failed', 
          message: error.message 
        });
      } else if (error.message.includes('timeout')) {
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
  @UseInterceptors(FileInterceptor('file'))
  async compressPdf(
    @UploadedFile() file: Express.Multer.File,
    @Query('quality') quality: string = 'moderate',
    @Req() req: Request,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No PDF file uploaded' });
      }
      
      // Security validation
      this.validateFileAndRequest(file, req, 'pdf');
      
      // Validate and sanitize quality parameter
      const sanitizedQuality = this.securityService.validateQualityParameter(quality);
      
      const sanitizedFilename = this.securityService.sanitizeFilename(file.originalname);
      this.logger.log(`Compressing PDF: ${sanitizedFilename}, quality: ${sanitizedQuality}, size: ${file.size} bytes`);
      
      const output = await this.appService.compressPdf(file, sanitizedQuality);
      
      res.set({ 
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="compressed_${sanitizedFilename}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY'
      });
      res.send(output);
    } catch (error) {
      this.logger.error('PDF compression error:', error.message);
      
      if (error.message.includes('Security validation failed')) {
        return res.status(400).json({ 
          error: 'Security validation failed', 
          message: error.message 
        });
      } else {
        res.status(500).json({ error: 'Failed to compress PDF', message: error.message });
      }
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

  // Security validation helper method
  private validateFileAndRequest(file: Express.Multer.File, req: Request, fileType: 'pdf' | 'office'): void {
    // Validate request source
    const sourceValidation = this.securityService.validateRequestSource(req);
    this.logger.log(`Request validation: ${sourceValidation.message}`);

    // Validate file security
    const fileValidation = this.securityService.validateFile(file, fileType);
    
    if (!fileValidation.isValid) {
      this.logger.error(`File validation failed for ${file.originalname}: ${fileValidation.errors.join(', ')}`);
      throw new Error(`Security validation failed: ${fileValidation.errors.join(', ')}`);
    }

    // Log warnings if any
    if (fileValidation.warnings.length > 0) {
      this.logger.warn(`File validation warnings for ${file.originalname}: ${fileValidation.warnings.join(', ')}`);
    }

    // Validate file buffer for malicious content
    const bufferValidation = this.securityService.validateFileBuffer(file.buffer);
    if (!bufferValidation.isValid) {
      this.logger.error(`Buffer validation failed for ${file.originalname}: ${bufferValidation.message}`);
      throw new Error(`File content validation failed: ${bufferValidation.message}`);
    }
  }
}