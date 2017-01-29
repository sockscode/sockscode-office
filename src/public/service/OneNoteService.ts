import { IOfficeService } from './OfficeService';
import { OfficeState } from '../reducers/Office'
import { Parser } from 'xml2js';
import { stripPrefix } from 'xml2js/lib/processors';

export class OneNoteService implements IOfficeService {
    constructor() {
        throw new Error('Not implemented');
    }

    onCodeChange(codeChangedListener: (data: string) => void) {
        throw new Error('Not implemented');
    }

    changeCode(code: string) {
        throw new Error('Not implemented');
    }
}