import { Controller, Post, UploadedFile, UseInterceptors, Res } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { AppService } from './app.service';
import * as multer from 'multer';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Post('convert-doc-to-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async convertDocToPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    const output = await this.appService.convertLibreOffice(file, 'pdf');
    res.set({ 'Content-Type': 'application/pdf' });
    res.send(output);
  }

  @Post('compress-pdf')
  @UseInterceptors(FileInterceptor('file'))
  async compressPdf(
    @UploadedFile() file: Express.Multer.File,
    @Res() res: Response,
  ) {
    const output = await this.appService.compressPdf(file);
    res.set({ 'Content-Type': 'application/pdf' });
    res.send(output);
  }
}