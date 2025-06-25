import { Controller, Post, Get, UploadedFile, UseInterceptors, Res, Body, Query } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { AppService } from './app.service';
import * as multer from 'multer';
import * as os from 'os';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  // Health check endpoint
  @Get()
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
  health(@Res() res: Response) {
    return res.status(200).json({ status: 'ok', timestamp: new Date().toISOString() });
  }

  // Word to PDF conversion
  @Post('convert-word-to-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async convertWordToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      const output = await this.appService.convertLibreOffice(file, 'pdf');
      res.set({ 'Content-Type': 'application/pdf' });
      res.send(output);
    } catch (error) {
      console.error('Word to PDF conversion error:', error.message);
      res.status(500).json({ error: 'Failed to convert Word to PDF', message: error.message });
    }
  }

  // Excel to PDF conversion
  @Post('convert-excel-to-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async convertExcelToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      const output = await this.appService.convertLibreOffice(file, 'pdf');
      res.set({ 'Content-Type': 'application/pdf' });
      res.send(output);
    } catch (error) {
      console.error('Excel to PDF conversion error:', error.message);
      res.status(500).json({ error: 'Failed to convert Excel to PDF', message: error.message });
    }
  }

  // PowerPoint to PDF conversion
  @Post('convert-ppt-to-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async convertPptToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }
      
      const output = await this.appService.convertLibreOffice(file, 'pdf');
      res.set({ 'Content-Type': 'application/pdf' });
      res.send(output);
    } catch (error) {
      console.error('PowerPoint to PDF conversion error:', error.message);
      res.status(500).json({ error: 'Failed to convert PowerPoint to PDF', message: error.message });
    }
  }

  // PDF to Word conversion
  @Post('convert-pdf-to-word')
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToWord(
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

      // Validate file size (PDFs larger than 50MB may cause issues)
      if (file.size > 50 * 1024 * 1024) {
        return res.status(400).json({ error: 'PDF file is too large. Please use a file smaller than 50MB.' });
      }
      
      console.log(`Converting PDF to Word: ${file.originalname}, size: ${file.size} bytes`);
      const output = await this.appService.convertLibreOffice(file, 'docx');
      
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
          message: 'This PDF contains complex formatting that cannot be converted. Try using a text-based PDF instead of a scanned document.' 
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
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToExcel(
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

      // Validate file size
      if (file.size > 50 * 1024 * 1024) {
        return res.status(400).json({ error: 'PDF file is too large. Please use a file smaller than 50MB.' });
      }
      
      console.log(`Converting PDF to Excel: ${file.originalname}, size: ${file.size} bytes`);
      const output = await this.appService.convertLibreOffice(file, 'xlsx');
      
      res.set({ 
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '.xlsx')}"` 
      });
      res.send(output);
    } catch (error) {
      console.error('PDF to Excel conversion error:', error.message);
      
      if (error.message.includes('timeout')) {
        return res.status(408).json({ 
          error: 'Conversion timeout', 
          message: 'The PDF conversion is taking too long. Please try with a smaller or simpler PDF.' 
        });
      } else if (error.message.includes('complex formatting')) {
        return res.status(422).json({ 
          error: 'Complex PDF format', 
          message: 'This PDF may not contain tabular data suitable for Excel conversion.' 
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
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToPpt(
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

      // Validate file size
      if (file.size > 50 * 1024 * 1024) {
        return res.status(400).json({ error: 'PDF file is too large. Please use a file smaller than 50MB.' });
      }
      
      console.log(`Converting PDF to PowerPoint: ${file.originalname}, size: ${file.size} bytes`);
      const output = await this.appService.convertLibreOffice(file, 'pptx');
      
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
          message: 'This PDF contains complex formatting that may not convert well to PowerPoint slides.' 
        });
      } else {
        return res.status(500).json({ 
          error: 'Failed to convert PDF to PowerPoint', 
          message: error.message 
        });
      }
    }
  }

  // PDF compression with quality options
  @Post('compress-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async compressPdf(
    @UploadedFile() file: Express.Multer.File,
    @Query('quality') quality: string = 'moderate',
    @Res() res: Response,
  ) {
    try {
      if (!file) {
        return res.status(400).json({ error: 'No PDF file uploaded' });
      }
      
      // Validate file type
      if (file.mimetype !== 'application/pdf') {
        return res.status(400).json({ error: 'Uploaded file is not a PDF' });
      }
      
      const output = await this.appService.compressPdf(file, quality);
      res.set({ 'Content-Type': 'application/pdf' });
      res.send(output);
    } catch (error) {
      console.error('PDF compression error:', error.message);
      res.status(500).json({ error: 'Failed to compress PDF', message: error.message });
    }
  }
}