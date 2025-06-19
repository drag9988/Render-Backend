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
    const output = await this.appService.convertLibreOffice(file, 'pdf');
    res.set({ 'Content-Type': 'application/pdf' });
    res.send(output);
  }

  // Excel to PDF conversion
  @Post('convert-excel-to-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async convertExcelToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    const output = await this.appService.convertLibreOffice(file, 'pdf');
    res.set({ 'Content-Type': 'application/pdf' });
    res.send(output);
  }

  // PowerPoint to PDF conversion
  @Post('convert-ppt-to-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async convertPptToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    const output = await this.appService.convertLibreOffice(file, 'pdf');
    res.set({ 'Content-Type': 'application/pdf' });
    res.send(output);
  }

  // PDF to Word conversion
  @Post('convert-pdf-to-word')
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToWord(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    const output = await this.appService.convertLibreOffice(file, 'docx');
    res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    res.send(output);
  }

  // PDF to Excel conversion
  @Post('convert-pdf-to-excel')
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToExcel(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    const output = await this.appService.convertLibreOffice(file, 'xlsx');
    res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    res.send(output);
  }

  // PDF to PowerPoint conversion
  @Post('convert-pdf-to-ppt')
  @UseInterceptors(FileInterceptor('file'))
  async convertPdfToPpt(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    const output = await this.appService.convertLibreOffice(file, 'pptx');
    res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    res.send(output);
  }

  // PDF compression with quality options
  @Post('compress-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async compressPdf(
    @UploadedFile() file: Express.Multer.File,
    @Query('quality') quality: string = 'moderate',
    @Res() res: Response,
  ) {
    const output = await this.appService.compressPdf(file, quality);
    res.set({ 'Content-Type': 'application/pdf' });
    res.send(output);
  }
}