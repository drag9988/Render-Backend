export declare class SecurityService {
    private readonly logger;
    constructor();
    validateRequest(request: any): boolean;
    sanitizeInput(input: string): string;
    isIpAllowed(ip: string): boolean;
}
