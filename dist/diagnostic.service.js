"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var DiagnosticService_1;
Object.defineProperty(exports, "__esModule", { value: true });
exports.DiagnosticService = void 0;
const common_1 = require("@nestjs/common");
let DiagnosticService = DiagnosticService_1 = class DiagnosticService {
    constructor() {
        this.logger = new common_1.Logger(DiagnosticService_1.name);
        this.logger.log('DiagnosticService initialized');
    }
    async getSystemHealth() {
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
    async runDiagnostics() {
        const health = await this.getSystemHealth();
        return Object.assign(Object.assign({}, health), { services: {
                fileValidation: true,
                onlyOffice: true,
                onlyOfficeEnhanced: true
            } });
    }
    logDiagnostics() {
        this.getSystemHealth().then(health => {
            this.logger.log(`System Health: ${JSON.stringify(health, null, 2)}`);
        });
    }
};
exports.DiagnosticService = DiagnosticService;
exports.DiagnosticService = DiagnosticService = DiagnosticService_1 = __decorate([
    (0, common_1.Injectable)(),
    __metadata("design:paramtypes", [])
], DiagnosticService);
//# sourceMappingURL=diagnostic.service.js.map