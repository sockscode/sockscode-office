import { OfficeState } from '../reducers/Office'

export class OfficeService {
    wordTextChangeEmitter: WordTextChangeEmitter;
    initialized: boolean;

    promiseInitialize(): Promise<OfficeState> {
        let resolver: (value?: OfficeState | PromiseLike<OfficeState>) => void;
        let promise = new Promise<OfficeState>((resolve, reject) => {
            resolver = resolve
        });
        Office.initialize = (reason) => {
            this.wordTextChangeEmitter = new WordTextChangeEmitter();
            this.initialized = true;
            resolver();
        }
        return promise;
    }

    onCodeChange(codeChangedListener: (data: string) => {}) {
        if (!this.initialized) {
            throw new Error('OfficeService should be initalized first');
        }
        this.wordTextChangeEmitter.subscribe('change', codeChangedListener);
    }

    changeCode(code: string) {
        if (!this.initialized) {
            throw new Error('OfficeService should be initalized first');
        }
        Word.run((context) => {
            context.document.body.insertText(code, 'Replace');
            return context.sync().then(function () {
                console.log('Replaced text with new code');
            });
        }).catch(function () {
            console.error("Failed to check for change of text", arguments);
        });
    }
}

type Listener = (data: any) => {};

/**
 * Listens for the text change every intervalTime and if the text changes => emits 'change' event
 */
class WordTextChangeEmitter {
    intervalTime: number;
    prevText: string;

    listeners: Map<String, Listener[]>;

    constructor(intervalTime = 1000) {
        this.listeners = new Map<String, Listener[]>();
        this.intervalTime = intervalTime;
        const changeChecker = () => {
            Word.run((context) => {
                const text = context.document.body.text;
                if (text != this.prevText) {
                    this.prevText = text;
                    this.emit('change', text);
                }
                return Promise.resolve();
            }).then(() => {
                changeChecker();
            }).catch(function () {
                console.error("Failed to check for change of text", arguments);
                changeChecker();
            });
        };
        setTimeout(changeChecker, intervalTime)
    }

    emit(event: string, data: any) {
        let listenersList = this.listeners.get(event);
        listenersList && listenersList.forEach((listener) => {
            listener(data);
        })
    }

    subscribe(event: string, eventListener: Listener) {
        let listenersList = this.listeners.get(event);
        if (!listenersList) {
            listenersList = [] as Listener[];
            this.listeners.set(event, listenersList)
        }
        listenersList.push(eventListener);
    }

    unsubscribe(event: string, eventListener: Listener): boolean {
        let listenersList = this.listeners.get(event);
        if (!listenersList) {
            return false;
        }
        let index = listenersList.indexOf(eventListener);
        if (~index) {
            listenersList.splice(index, 1);
            return true;
        }
        return false;
    }
}