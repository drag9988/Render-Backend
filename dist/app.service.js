"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.AppService = void 0;
const common_1 = require("@nestjs/common");
const fs = require("fs/promises");
const child_process_1 = require("child_process");
const util_1 = require("util");
const execAsync = (0, util_1.promisify)(child_process_1.exec);
let AppService = class AppService {
    async convertLibreOffice(file, format) {
        const timestamp = Date.now();
        const tempInput = /tmp/$, { timestamp }, _$, { file, originalname };
        const tempOutput = tempInput.replace(/\.[^.]+$/, $, { format });
        await fs.writeFile(tempInput, file.buffer);
        await execAsync(libreoffice--, headless--, convert - to, $, { format }--, outdir / tmp, $, { tempInput });
        const result = await fs.readFile(tempOutput);
        await fs.unlink(tempInput);
        await fs.unlink(tempOutput);
        return result;
    }
    async compressPdf(file) {
        const timestamp = Date.now();
        const input = /tmp/$, { timestamp }, _input, pdf;
        const output = /tmp/$, { timestamp }, _output, pdf;
        await fs.writeFile(input, file.buffer);
        await execAsync(gs - sDEVICE, pdfwrite - dCompatibilityLevel, 1.4 - dPDFSETTINGS, /ebook -dNOPAUSE -dQUIET -dBATCH -sOutputFile=${output} ${input});
        const result = await fs.readFile(output);
        await fs.unlink(input);
        await fs.unlink(output);
        return result;
    }
};
exports.AppService = AppService;
exports.AppService = AppService = __decorate([
    (0, common_1.Injectable)()
], AppService);
//# sourceMappingURL=app.service.js.map