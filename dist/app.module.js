"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.AppModule = void 0;
const common_1 = require("@nestjs/common");
const app_controller_1 = require("./app.controller");
const app_service_1 = require("./app.service");
const onlyoffice_service_1 = require("./onlyoffice.service");
const onlyoffice_enhanced_service_1 = require("./onlyoffice-enhanced.service");
const file_validation_service_1 = require("./file-validation.service");
const security_service_1 = require("./security.service");
const diagnostic_service_1 = require("./diagnostic.service");
const throttler_1 = require("@nestjs/throttler");
let AppModule = class AppModule {
};
exports.AppModule = AppModule;
exports.AppModule = AppModule = __decorate([
    (0, common_1.Module)({
        imports: [
            throttler_1.ThrottlerModule.forRoot([{
                    ttl: 60000,
                    limit: 100,
                }]),
        ],
        controllers: [app_controller_1.AppController],
        providers: [
            app_service_1.AppService,
            onlyoffice_service_1.OnlyOfficeService,
            onlyoffice_enhanced_service_1.OnlyOfficeEnhancedService,
            file_validation_service_1.FileValidationService,
            security_service_1.SecurityService,
            diagnostic_service_1.DiagnosticService,
        ],
    })
], AppModule);
//# sourceMappingURL=app.module.js.map