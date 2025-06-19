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

    // Create a unique filename with timestamp
    const timestamp = Date.now();
    const sanitizedFilename = file.originalname.replace(/[^a-zA-Z0-9.]/g, '_');
    
    // Use the OS temp directory or fallback to /tmp
    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
    console.log(`Using temp directory: ${tempDir}`);
    
    const tempInput = `${tempDir}/${timestamp}_${sanitizedFilename}`;
    const tempOutput = tempInput.replace(/\.[^.]+$/, `.${format}`);

    try {
      console.log(`Starting conversion from ${file.mimetype} to ${format}`);
      
      // Write the uploaded file to disk
      await fs.writeFile(tempInput, file.buffer);
      console.log(`File written to ${tempInput}`);
      
      // Execute LibreOffice conversion
      const command = `libreoffice --headless --convert-to ${format} --outdir ${tempDir} ${tempInput}`;
      console.log(`Executing command: ${command}`);
      
      const { stdout, stderr } = await this.execAsync(command);
      
      if (stdout) {
        console.log(`LibreOffice conversion output: ${stdout}`);
      }
      
      if (stderr) {
        console.error(`LibreOffice conversion error: ${stderr}`);
      }

      // Check if output file exists
      try {
        await fs.access(tempOutput);
      } catch (err) {
        console.error(`Output file not found at ${tempOutput}`);
        throw new Error(`Conversion failed: Output file not created`);
      }

      // Read the converted file
      console.log(`Reading converted file from ${tempOutput}`);
      const result = await fs.readFile(tempOutput);
      console.log(`Successfully read converted file, size: ${result.length} bytes`);
      return result;
    } catch (error) {
      console.error(`File conversion error: ${error.message}`);
      throw new Error(`Failed to convert document to ${format}: ${error.message}`);
    } finally {
      // Clean up temporary files
      try {
        console.log(`Cleaning up temporary files`);
        await fs.unlink(tempInput).catch((err) => console.error(`Failed to delete input file: ${err.message}`));
        await fs.unlink(tempOutput).catch((err) => console.error(`Failed to delete output file: ${err.message}`));
      } catch (cleanupError) {
        console.error(`Cleanup error: ${cleanupError.message}`);
      }
    }
  }

  async compressPdf(file: Express.Multer.File, quality: string = 'moderate'): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid PDF file provided');
    }

    // Create a unique filename with timestamp
    const timestamp = Date.now();
    
    // Use the OS temp directory or fallback to /tmp
    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
    console.log(`Using temp directory for PDF compression: ${tempDir}`);
    
    const input = `${tempDir}/${timestamp}_input.pdf`;
    const output = `${tempDir}/${timestamp}_output.pdf`;
    
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
    console.log(`PDF compression quality: ${quality}, using setting: ${pdfSetting}`);

    try {
      console.log(`Starting PDF compression, input size: ${file.buffer.length} bytes`);
      
      // Write the uploaded PDF to disk
      await fs.writeFile(input, file.buffer);
      console.log(`PDF written to ${input}`);

      // Execute Ghostscript for PDF compression
      const command = `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=${pdfSetting} -dNOPAUSE -dQUIET -dBATCH -sOutputFile=${output} ${input}`;
      console.log(`Executing command: ${command}`);
      
      const { stdout, stderr } = await this.execAsync(command);
      
      if (stdout) {
        console.log(`Ghostscript output: ${stdout}`);
      }
      
      if (stderr) {
        console.error(`Ghostscript compression error: ${stderr}`);
      }

      // Check if output file exists
      try {
        await fs.access(output);
      } catch (err) {
        console.error(`Compressed PDF not found at ${output}`);
        throw new Error(`Compression failed: Output file not created`);
      }

      // Read the compressed PDF
      console.log(`Reading compressed PDF from ${output}`);
      const result = await fs.readFile(output);
      console.log(`Successfully compressed PDF, original size: ${file.buffer.length} bytes, compressed size: ${result.length} bytes`);
      
      // If the compressed file is larger than the original, return the original
      if (result.length > file.buffer.length) {
        console.log(`Compressed file is larger than original, returning original file`);
        return file.buffer;
      }
      
      return result;
    } catch (error) {
      console.error(`PDF compression error: ${error.message}`);
      throw new Error(`Failed to compress PDF: ${error.message}`);
    } finally {
      // Clean up temporary files
      try {
        console.log(`Cleaning up temporary PDF files`);
        await fs.unlink(input).catch((err) => console.error(`Failed to delete input PDF: ${err.message}`));
        await fs.unlink(output).catch((err) => console.error(`Failed to delete output PDF: ${err.message}`));
      } catch (cleanupError) {
        console.error(`Cleanup error: ${cleanupError.message}`);
      }
    }
  }
}