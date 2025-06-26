import { Injectable, Logger } from '@nestjs/common';
import { exec } from 'child_process';
import { promisify } from 'util';

@Injectable()
export class DiagnosticService {
  private readonly logger = new Logger(DiagnosticService.name);
  private readonly execAsync = promisify(exec);

  async checkSystemRequirements(): Promise<{
    libreoffice: boolean;
    ghostscript: boolean;
    pdftools: boolean;
    system: any;
  }> {
    const results = {
      libreoffice: false,
      ghostscript: false,
      pdftools: false,
      system: {
        platform: process.platform,
        arch: process.arch,
        memory: process.memoryUsage(),
        uptime: process.uptime()
      }
    };

    try {
      // Check LibreOffice
      await this.execAsync('libreoffice --version', { timeout: 10000 });
      results.libreoffice = true;
      this.logger.log('✅ LibreOffice found');
    } catch (error) {
      this.logger.error('❌ LibreOffice not found:', error.message);
      
      // Try alternative commands
      try {
        await this.execAsync('soffice --version', { timeout: 10000 });
        results.libreoffice = true;
        this.logger.log('✅ LibreOffice found via soffice command');
      } catch (altError) {
        this.logger.error('❌ soffice command also failed:', altError.message);
      }
    }

    try {
      // Check Ghostscript
      await this.execAsync('gs --version', { timeout: 10000 });
      results.ghostscript = true;
      this.logger.log('✅ Ghostscript found');
    } catch (error) {
      this.logger.error('❌ Ghostscript not found:', error.message);
    }

    try {
      // Check PDF tools
      await this.execAsync('pdfinfo -v', { timeout: 10000 });
      results.pdftools = true;
      this.logger.log('✅ PDF tools found');
    } catch (error) {
      this.logger.error('❌ PDF tools not found:', error.message);
    }

    return results;
  }

  async testLibreOfficeConversion(): Promise<boolean> {
    try {
      const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
      
      // Create a simple test document
      const testContent = 'This is a test document for LibreOffice conversion.';
      const testFile = `${tempDir}/test_${Date.now()}.txt`;
      const outputFile = `${tempDir}/test_${Date.now()}.pdf`;
      
      const fs = require('fs').promises;
      await fs.writeFile(testFile, testContent);
      
      // Try to convert
      const command = `libreoffice --headless --convert-to pdf --outdir ${tempDir} ${testFile}`;
      this.logger.log(`Testing LibreOffice with: ${command}`);
      
      const { stdout, stderr } = await this.execAsync(command, { timeout: 30000 });
      
      if (stdout) this.logger.log(`Test output: ${stdout}`);
      if (stderr) this.logger.error(`Test error: ${stderr}`);
      
      // Check if output was created
      const expectedOutput = testFile.replace('.txt', '.pdf');
      try {
        await fs.access(expectedOutput);
        this.logger.log('✅ LibreOffice test conversion successful');
        
        // Clean up
        await fs.unlink(testFile).catch(() => {});
        await fs.unlink(expectedOutput).catch(() => {});
        
        return true;
      } catch (accessError) {
        this.logger.error('❌ LibreOffice test conversion failed - no output file');
        return false;
      }
      
    } catch (error) {
      this.logger.error('❌ LibreOffice test conversion failed:', error.message);
      return false;
    }
  }
}
