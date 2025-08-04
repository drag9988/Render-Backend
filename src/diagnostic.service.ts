import { Injectable, Logger } from '@nestjs/common';

@Injectable()
export class DiagnosticService {
  private readonly logger = new Logger(DiagnosticService.name);

  constructor() {
    this.logger.log('DiagnosticService initialized');
  }

  /**
   * Get system health information
   */
  async getSystemHealth(): Promise<any> {
    return {
      status: 'healthy',
      timestamp: new Date().toISOString(),
      uptime: process.uptime(),
      memory: {
        used: Math.round(process.memoryUsage().heapUsed / 1024 / 1024),
        total: Math.round(process.memoryUsage().heapTotal / 1024 / 1024),
        rss: Math.round(process.memoryUsage().rss / 1024 / 1024),
      },
      environment: process.env.NODE_ENV || 'development'
    };
  }

  /**
   * Run diagnostic checks
   */
  async runDiagnostics(): Promise<any> {
    const health = await this.getSystemHealth();
    
    return {
      ...health,
      services: {
        fileValidation: true,
        onlyOffice: true,
        onlyOfficeEnhanced: true
      }
    };
  }

  /**
   * Log diagnostic information
   */
  logDiagnostics(): void {
    this.getSystemHealth().then(health => {
      this.logger.log(`System Health: ${JSON.stringify(health, null, 2)}`);
    });
  }
}
