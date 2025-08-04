export declare class DiagnosticService {
    private readonly logger;
    constructor();
    getSystemHealth(): Promise<any>;
    runDiagnostics(): Promise<any>;
    logDiagnostics(): void;
}
