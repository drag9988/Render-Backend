import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { OnlyOfficeService } from './onlyoffice.service';
import { OnlyOfficeEnhancedService } from './onlyoffice-enhanced.service';
import { FileValidationService } from './file-validation.service';
import { SecurityService } from './security.service';
import { DiagnosticService } from './diagnostic.service';
import { ThrottlerModule } from '@nestjs/throttler';

@Module({
  imports: [
    ThrottlerModule.forRoot([{
      ttl: 60000,
      limit: 100,
    }]),
  ],
  controllers: [AppController],
  providers: [
    AppService,
    OnlyOfficeService,
    OnlyOfficeEnhancedService,
    FileValidationService,
    SecurityService,
    DiagnosticService,
  ],
})
export class AppModule {}