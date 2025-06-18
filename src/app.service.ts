import { Injectable } from '@nestjs/common';
import * as fs from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

@Injectable()
export class AppService {
  async convertLibreOffice(file: Express.Multer.File, format: string): Promise<Buffer> {
    const tempInput = `/tmp/${Date.now()}_${file.originalname}`;
    const tempOutput = tempInput.replace(/\.[^.]+$/, \`.${format}\`);

    await fs.writeFile(tempInput, file.buffer);
    await execAsync(\`libreoffice --headless --convert-to ${format} --outdir /tmp ${tempInput}\`);

    const result = await fs.readFile(tempOutput);
    await fs.unlink(tempInput);
    await fs.unlink(tempOutput);
    return result;
  }

  async compressPdf(file: Express.Multer.File): Promise<Buffer> {
    const input = `/tmp/${Date.now()}_input.pdf`;
    const output = input.replace('input', 'output');

    await fs.writeFile(input, file.buffer);
    await execAsync(
      \`gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/ebook -dNOPAUSE -dQUIET -dBATCH -sOutputFile=${output} ${input}\`,
    );

    const result = await fs.readFile(output);
    await fs.unlink(input);
    await fs.unlink(output);
    return result;
  }
}