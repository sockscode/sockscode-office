import { WordService } from './WordService';
import { OneNoteService } from './OneNoteService';

export interface IOfficeService {
    onCodeChange: (codeChangedListener: (data: string) => void) => void;
    changeCode: (code: string) => void;
}

export class OfficeService {
    initialized: boolean;
    officeService: IOfficeService;

    promiseInitialize(): Promise<void> {
        if (this.initialized) {
            return Promise.resolve();
        }
        let resolver: () => void;
        let rejecter: (reason?: any) => void;
        let promise = new Promise<void>((resolve, reject) => {
            resolver = resolve
            rejecter = reject;
        });
        Office.initialize = (reason) => {
            this.initialized = true;
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                this.initialized = true;
                this.officeService = new WordService();
                resolver();
            } else if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1)) {
                this.initialized = true;
                this.officeService = new OneNoteService();
                resolver();
            } else {
                this.initialized = false;
                rejecter('Your application is not supported. Only Word and OneNote are supported.');
            }
        }//fixme catch?
        return promise;
    }

    onCodeChange(codeChangedListener: (data: string) => void) {
        if (!this.initialized) {
            throw new Error('OfficeService should be initalized first');
        }
        this.officeService.onCodeChange(codeChangedListener);
    }

    changeCode(code: string) {
        if (!this.initialized) {
            throw new Error('OfficeService should be initalized first');
        }
        this.officeService.changeCode(code);
    }
}