import { OfficeState } from '../reducers/Office'
import { Parser } from 'xml2js';
import { stripPrefix } from 'xml2js/lib/processors';

export class OfficeService {
    // Using the WordJS API. Creates content control. Sets the tag and title property
    // for content control. 
    //@deprecated office online doesn't support this the right way right now.
    public static addContentControls() {
        const deleteOldContentControls = Word.run((ctx) => {
            const ccs = ctx.document.contentControls.getByTag("sockscode");
            ctx.load(ccs);
            return ctx.sync().then(() => {
                const items = ccs.items.map((item) => {
                    item.cannotDelete = false;
                    ctx.load(item);
                    return item;
                });

                return ctx.sync().then(() => {
                    items.forEach((item) => {
                        item.delete(true);
                    })
                    return ctx.sync();
                });
            });
        })

        return deleteOldContentControls.then(() => {
            return Word.run((ctx) => {
                const range = ctx.document.body.insertText('//code goes here', 'Start');
                const controll = range.insertContentControl();
                controll.cannotEdit = false;
                controll.cannotDelete = false;
                controll.tag = 'sockscode';
                controll.title = 'sockscode';
                return ctx.sync();
            });
        });
    }

    wordTextChangeEmitter: WordTextChangeEmitter;
    initialized: boolean;
    changeCodeBuffered: (code: string) => void;

    constructor() {
        let timeoutId = 0;
        this.changeCodeBuffered = (code: string) => {
            clearTimeout(timeoutId);
            timeoutId = (setTimeout(() => {
                this.changeCode(code);
            }, 200) as any) as number;
        };
    }

    promiseInitialize(): Promise<void> {
        if (this.initialized) {
            return Promise.resolve();
        }
        let resolver: () => void;
        let promise = new Promise<void>((resolve, reject) => {
            resolver = resolve
        });
        Office.initialize = (reason) => {
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                this.wordTextChangeEmitter = new WordTextChangeEmitter();
                this.initialized = true;
                resolver();
            } else {
                //fixme reject 
            }
        }//fixme catch?
        return promise;
    }

    onCodeChange(codeChangedListener: (data: string) => void) {
        if (!this.initialized) {
            throw new Error('OfficeService should be initalized first');
        }
        this.wordTextChangeEmitter.subscribe('change', codeChangedListener);
    }

    changeCode(code: string) {
        if (!this.initialized) {
            throw new Error('OfficeService should be initalized first');
        }
        this.wordTextChangeEmitter.suspendEventRecording();
        Word.run((ctx) => {
            ctx.document.body.insertText(code, 'Replace');
            return ctx.sync().then(() => {
                this.wordTextChangeEmitter.prevText = code;
                const html = ctx.document.body.getHtml();
                return ctx.sync().then(() => {
                    this.wordTextChangeEmitter.prevText = HtmlParser.parseHtml(html);
                    this.wordTextChangeEmitter.resumeEventRecording();
                });
            });
        }).catch(function () {
            this.wordTextChangeEmitter.resumeEventRecording();
            console.error("Failed to check for change of text", arguments);
        }.bind(this));
    }
}

type Listener = (data: any) => void;

/**
 * Listens for the text change every intervalTime and if the text changes => emits 'change' event
 */
class WordTextChangeEmitter {
    public prevText: string;
    intervalTime: number;
    _suspendEventRecording = 0;

    listeners: Map<String, Listener[]>;

    constructor(intervalTime = 500) {
        this.listeners = new Map<String, Listener[]>();
        this.intervalTime = intervalTime;

        const changeChecker = () => {
            if (this._suspendEventRecording) {
                setTimeout(changeChecker, intervalTime);
                return;
            }
            let start = new Date();
            let elapsed = (event: string) => {
                console.log(event, (new Date() as any) - (start as any));
            }
            Word.run((ctx) => {
                elapsed("Started");
                const html = ctx.document.body.getHtml();
                return ctx.sync().then(() => {
                    elapsed("html");
                    return HtmlParser.parseHtml(html);
                });
            }).then((text: string) => {
                elapsed("Parsed");
                setTimeout(changeChecker, intervalTime);
                if (text !== this.prevText && !this._suspendEventRecording) {
                    console.log(text);
                    this.prevText = text;
                    this.emit('change', text);
                }
            }).catch(() => {
                setTimeout(changeChecker, intervalTime);
            });
        }
        setTimeout(changeChecker, intervalTime);

        /*const changeChecker = () => {
            let start = new Date();
            let elapsed = (event: string) => {
                console.log(event, (new Date() as any) - (start as any));
            }
            Word.run((ctx) => {
                elapsed("Started");
                const paragraphs = ctx.document.body.paragraphs;
                ctx.load(paragraphs);
                return ctx.sync().then(() => {
                    elapsed("Paragraps");
                    const paragraph = paragraphs.items[0];
                    if (paragraph) {
                        const html = paragraph.getHtml();
                        return ctx.sync().then(() => {
                            elapsed("ooxml");
                            console.log(html);
                            return Promise.resolve(html.value);
                            //return OoxmlTextParser.parseOoxm(html);
                        });
                    }
                    return Promise.reject('no paragraph');
                });
            }).then((text: string) => {
                elapsed("Parsed");
                setTimeout(changeChecker, intervalTime);
                if (text !== this.prevText) {
                    this.prevText = text;
                    this.emit('change', text);
                }
            }).catch(() => {
                setTimeout(changeChecker, intervalTime);
            });
        }
        setTimeout(changeChecker, intervalTime);*/

        /*const changeChecker = () => {
            let start = new Date();
            let elapsed = (event: string) => {
                console.log(event, (new Date() as any) - (start as any));
            }
            Word.run((ctx) => {
                elapsed("Started");
                const paragraphs = ctx.document.body.paragraphs;
                ctx.load(paragraphs);
                return ctx.sync().then(() => {
                    elapsed("Paragraps");
                    const paragraph = paragraphs.items[0];
                    if (paragraph) {
                        const ooxml = paragraph.getOoxml();;
                        return ctx.sync().then(() => {
                            elapsed("ooxml");
                            return OoxmlTextParser.parseOoxm(ooxml);
                        });
                    }
                    return Promise.reject('no paragraph');
                });
            }).then((text: string) => {
                elapsed("Parsed");
                setTimeout(changeChecker, intervalTime);
                if (text !== this.prevText) {
                    this.prevText = text;
                    this.emit('change', text);
                }
            }).catch(() => {
                setTimeout(changeChecker, intervalTime);
            });
        }
        setTimeout(changeChecker, intervalTime);*/

        //content controlls based (not working in online office)
        /*this.listeners = new Map<String, Listener[]>();
        this.intervalTime = intervalTime;
        const changeChecker = () => {
            Word.run((ctx) => {
                const ccs = ctx.document.contentControls.getByTag("sockscode");
                ctx.load(ccs);
                return ctx.sync().then(() => {
                    const contentControll = ccs.items[0];
                    //document.body.innerHTML = contentControll.getTextRanges + '';
                    if (contentControll) {
                        const textRanges = contentControll.getTextRanges(['\n']);
                        ctx.load(textRanges);
                        return ctx.sync().then(() => {
                            textRanges.items.forEach((textRange) => {
                                //document.body.innerHTML = textRange.text;
                                console.log(textRange.text);
                            });
                        });
                    }
                    return Promise.reject('no contentControll');
                });
            }).then(() => {
                setTimeout(changeChecker, intervalTime);
            }).catch(function (error) {
                //document.body.innerHTML = JSON.stringify(arguments);
                setTimeout(changeChecker, intervalTime);
            });
        }
        setTimeout(changeChecker, intervalTime);*/

        // const changeChecker = () => {
        //     // Word.run((context) => {
        //     //     // Create a proxy object for the document body.                
        //     //     const body = context.document.body;
        //     //     const ooxml = body.getOoxml();
        //     //     // Queue a commmand to load the text in document body.                
        //     //     return context.sync().then(() => {
        //     //         return OoxmlTextParser.parseOoxm(ooxml).then((text: string) => {
        //     //             if (text !== this.prevText) {
        //     //                 this.prevText = text;
        //     //                 this.emit('change', text);
        //     //             }
        //     //         })
        //     //     });
        //     // }).then(() => {
        //     //     setTimeout(changeChecker, intervalTime);
        //     // }).catch(function () {
        //     //     console.error("Failed to check for change of text", arguments);
        //     //     setTimeout(changeChecker, intervalTime);
        //     // });

        //     // Run a batch operation against the Word object model.
        //     Word.run(function (context) {

        //         // Create a proxy object for the paragraphs collection.
        //         const paragraphs = context.document.body.paragraphs;

        //         context.load(paragraphs, { top: 1 });

        //         // Synchronize the document state by executing the queued commands, 
        //         // and return a promise to indicate task completion.
        //         return context.sync().then(function () {

        //             // Queue a command to get the last paragraph and create a 
        //             // proxy paragraph object.
        //             const paragraph = paragraphs.items[0];
        //             const ooxml = paragraph.getRange()

        //             // Synchronize the document state by executing the queued commands, 
        //             // and return a promise to indicate task completion.
        //             return context.sync().then(function () {
        //                 console.log('Selected the last paragraph.');
        //             });
        //         });
        //     })
        //         .catch(function (error) {
        //             console.log('Error: ' + JSON.stringify(error));
        //             if (error instanceof OfficeExtension.Error) {
        //                 console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        //             }
        //         });

        // };
        // setTimeout(changeChecker, intervalTime)
    }

    suspendEventRecording() {
        this._suspendEventRecording++;
    }

    resumeEventRecording() {
        this._suspendEventRecording = Math.max(0, this._suspendEventRecording - 1);
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



/**
 * Ooxml retrieving is stupidly slow, 
 */


interface OoxmlXml2jsObject {
    "pkg:package": {
        "pkg:part": [{
            "$": {
                "pkg:name": string,
            }
            "pkg:xmlData": [
                {
                    "w:document"?: [{
                        "w:body": [ //Document Body
                            {
                                "w:p": [ //Paragraph
                                    {
                                        "w:r": [ //Phonetic Guide Text Run http://www.datypic.com/sc/ooxml/e-w_r-1.html
                                            {
                                                "w:br": [
                                                    string
                                                ],
                                                "w:t": [
                                                    string //text here
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    }]
                }
            ]
        }]
    }
}

/**
 * 
 * Parser for html word format to get text with line breaks. We need to use this, because context.document.body.text doesn't have any line breaks inside:(
 */
class HtmlParser {
    public static parseHtml(html: { value: string }): string {
        const documnetFragment = document.createDocumentFragment();
        const htmlEl = document.createElement('html');
        htmlEl.innerHTML = html.value;
        documnetFragment.appendChild(htmlEl);
        return Array.from(documnetFragment.querySelectorAll('.Paragraph')).map((p) => {
            return Array.from(p.querySelectorAll('.TextRun, .LineBreakBlob')).map((el) => {
                if (el.classList.contains('LineBreakBlob')) {
                    return '\n';
                }
                return el.textContent;
            }).join('');
        }).join('\n');
    }
}

/**
 * 
 * Parser for ooxml word format to get text with line breaks. We need to use this, because context.document.body.text doesn't have any line breaks inside:(
 * @deprecated because ooxml retrieving is to slow
 */
class OoxmlTextParser {
    public static parseOoxm(ooxml: { value: string }): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            const parser = new Parser();
            parser.parseString(ooxml.value, (err: any, result: OoxmlXml2jsObject) => {
                if (err) {
                    reject(err);
                    return;
                }

                if (result['pkg:package'] && result['pkg:package']['pkg:part']) {
                    const pkgPart = result['pkg:package']['pkg:part'];
                    const documentPart = pkgPart.find((part) => {
                        return part['$'] && part['$']['pkg:name'] === '/word/document.xml';
                    });
                    if (documentPart && documentPart['pkg:xmlData'] && documentPart['pkg:xmlData'][0] && documentPart['pkg:xmlData'][0]['w:document'] &&
                        documentPart['pkg:xmlData'][0]['w:document'][0] && documentPart['pkg:xmlData'][0]['w:document'][0]['w:body']) {
                        const body = documentPart['pkg:xmlData'][0]['w:document'][0]['w:body'];
                        let textContent = '';
                        let firstP = true;
                        body.forEach((paragraphSuposedly, i) => {
                            if (paragraphSuposedly['w:p']) {
                                if (!firstP) {
                                    textContent += '\n';
                                }
                                const texts = paragraphSuposedly['w:p'];
                                let firstR = true;
                                texts && texts.forEach((text) => {
                                    if (!firstR) {
                                        textContent += '\n';
                                    }
                                    text['w:r'] && text['w:r'].forEach((wr) => {
                                        if (wr['w:br']) {
                                            textContent += '\n';
                                        }
                                        if (wr['w:t']) {
                                            textContent += wr['w:t'];
                                        }

                                    });
                                    firstR = false;
                                });
                                firstP = false;
                            } //else just skip. we cannot use any other type right now                            
                        });
                        resolve(textContent);
                    } else {
                        reject('malformed ooxml. No document part inside.' + JSON.stringify(result));
                        return;
                    }
                } else {
                    reject('malformed ooxml.' + JSON.stringify(result));
                    return;
                }
            });
        });
    }
}