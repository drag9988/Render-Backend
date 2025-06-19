import { Injectable } from '@nestjs/common';
import * as fs from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as multer from 'multer';

@Injectable()
export class AppService {
  private readonly execAsync = promisify(exec);

  async convertLibreOffice(file: Express.Multer.File, format: string): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid file provided');
    }

    const timestamp = Date.now();
    const tempInput = `/tmp/${timestamp}_${file.originalname}`;
    const tempOutput = tempInput.replace(/\.[^.]+$/, `.${format}`);

    try {
      // Write the uploaded file to disk
      await fs.writeFile(tempInput, file.buffer);
      
      // Execute LibreOffice conversion
      const { stdout, stderr } = await this.execAsync(
        `libreoffice --headless --convert-to ${format} --outdir /tmp ${tempInput}`,
      );
      
      if (stderr && !stderr.includes('convert')) {
        console.error(`LibreOffice conversion error: ${stderr}`);
      }

      // Read the converted file
      const result = await fs.readFile(tempOutput);
      return result;
    } catch (error) {
      console.error(`File conversion error: ${error.message}`);
      throw new Error(`Failed to convert document to ${format}: ${error.message}`);
    } finally {
      // Clean up temporary files
      try {
        await fs.unlink(tempInput).catch(() => {});
        await fs.unlink(tempOutput).catch(() => {});
      } catch (cleanupError) {
        console.error(`Cleanup error: ${cleanupError.message}`);
      }
    }
  }

  async compressPdf(file: Express.Multer.File, quality: string = 'moderate'): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid PDF file provided');
    }

    const timestamp = Date.now();
    const input = `/tmp/${timestamp}_input.pdf`;
    const output = `/tmp/${timestamp}_output.pdf`;
    
    // Map quality settings to Ghostscript PDF settings
    // /screen - low quality, smaller size (72 dpi)
    // /ebook - medium quality, medium size (150 dpi)
    // /printer - high quality, larger size (300 dpi)
    const qualitySettings = {
      'low': '/screen',
      'moderate': '/ebook',
      'high': '/printer'
    };
    
    // Default to moderate if invalid quality is provided
    const pdfSetting = qualitySettings[quality] || '/ebook';

    try {
      // Write the uploaded PDF to disk
      await fs.writeFile(input, file.buffer);

      // Execute Ghostscript for PDF compression
      const { stdout, stderr } = await this.execAsync(
        `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=${pdfSetting} -dNOPAUSE -dQUIET -dBATCH -sOutputFile=${output} ${input}`,
      );
      
      if (stderr) {
        console.error(`Ghostscript compression error: ${stderr}`);
      }

      // Read the compressed PDF
      const result = await fs.readFile(output);
      return result;
    } catch (error) {
      console.error(`PDF compression error: ${error.message}`);
      throw new Error(`Failed to compress PDF: ${error.message}`);
    } finally {
      // Clean up temporary files
      try {
        await fs.unlink(input).catch(() => {});
        await fs.unlink(output).catch(() => {});
      } catch (cleanupError) {
        console.error(`Cleanup error: ${cleanupError.message}`);
      }
    }
  }
}