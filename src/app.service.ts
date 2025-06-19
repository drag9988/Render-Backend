import { Injectable } from '@nestjs/common';
import * as fs from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as multer from 'multer';

@Injectable()
export class AppService {
  private readonly execAsync = promisify(exec);

  async convertLibreOffice(file: Express.Multer.File, format: string): Promise<Buffer> {
    const timestamp = Date.now();
    const tempInput = `/tmp/${timestamp}_${file.originalname}`;
    const tempOutput = tempInput.replace(/\.[^.]+$/, `.${format}`);

    await fs.writeFile(tempInput, file.buffer);

    await this.execAsync(
      `libreoffice --headless --convert-to ${format} --outdir /tmp ${tempInput}`,
    );

    const result = await fs.readFile(tempOutput);
    await fs.unlink(tempInput);
    await fs.unlink(tempOutput);

    return result;
  }

  async compressPdf(file: Express.Multer.File, quality: string = 'moderate'): Promise<Buffer> {
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

    await fs.writeFile(input, file.buffer);

    await this.execAsync(
      `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=${pdfSetting} -dNOPAUSE -dQUIET -dBATCH -sOutputFile=${output} ${input}`,
    );

    const result = await fs.readFile(output);
    await fs.unlink(input);
    await fs.unlink(output);

    return result;
  }
}