interface DealAssumptions {
    dealName: string;
    dealSize: number;
    ltv: number;
    holdingPeriod: number;
    revenueGrowth: number;
    exitMultiple: number;
    selectedRange?: string;
    rangeData?: any[][];
}
interface ChatMessage {
    role: 'user' | 'assistant';
    content: string;
}
declare class MAModelingAddin {
    private chatMessages;
    private selectedRange;
    private uploadedFiles;
    constructor();
    private initializeAddin;
    private initializeFileUpload;
    private handleFileSelection;
    private updateFileDisplay;
    private removeFile;
    private formatFileSize;
    private selectAssumptionsRange;
    private parseAssumptionsFromRange;
    private updateFormWithAssumptions;
    private collectAssumptions;
    private generateModel;
    private createModelSheets;
    private populateAssumptionsSheet;
    private populateNPVSheet;
    private populatePLSheet;
    private populateCFSheet;
    private calculateMetrics;
    private generateCashFlows;
    private calculateMetricsFallback;
    private validateModel;
    private sendChatMessage;
    private processWithAI;
    private processUploadedFiles;
    private readTextFile;
    private getExcelContext;
    private offerToImplementFormula;
    private implementSuggestedFormula;
    private addChatMessage;
    private showLoading;
    private showStatus;
    private executeCommand;
    private setValueCommand;
    private addToCellCommand;
    private setFormulaCommand;
    private formatCellCommand;
    private generateAssumptionsTemplate;
    private fillAssumptionsData;
    private getColumnLetter;
}
