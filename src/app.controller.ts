import { Controller, Post, UploadedFile, UseInterceptors, Res, Body, Query } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { AppService } from './app.service';
import * as multer from 'multer';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

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
      
      const output = await this.appService.convertLibreOffice(file, 'docx');
      res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      res.send(output);
    } catch (error) {
      console.error('PDF to Word conversion error:', error.message);
      res.status(500).json({ error: 'Failed to convert PDF to Word', message: error.message });
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
      
      const output = await this.appService.convertLibreOffice(file, 'xlsx');
      res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      res.send(output);
    } catch (error) {
      console.error('PDF to Excel conversion error:', error.message);
      res.status(500).json({ error: 'Failed to convert PDF to Excel', message: error.message });
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
      
      const output = await this.appService.convertLibreOffice(file, 'pptx');
      res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      res.send(output);
    } catch (error) {
      console.error('PDF to PowerPoint conversion error:', error.message);
      res.status(500).json({ error: 'Failed to convert PDF to PowerPoint', message: error.message });
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