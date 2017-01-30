import { IOfficeService } from './OfficeService';
import { OfficeState } from '../reducers/Office'
import { Emitter } from './emitter/Emitter';
import { Parser } from 'xml2js';
import { stripPrefix } from 'xml2js/lib/processors';

export class WordService implements IOfficeService {
    private _wordTextChangeEmitter: WordTextChangeEmitter;
    private _changingCode: boolean;

    constructor() {
        let timeoutId = 0;
        const changeCode = this.changeCode.bind(this);
        /**
         * We buffer execution because it's too slow to change the text on every change.
         * We'll be able to move away from this when 'small changes' will be introduced.
         */
        this.changeCode = (code: string) => {
            clearTimeout(timeoutId);
            timeoutId = (setTimeout(() => {
                if (this._changingCode) {
                    changeCode(code);
                    return;
                }
                changeCode(code);
            }, 200) as any) as number;
        };
        this._wordTextChangeEmitter = new WordTextChangeEmitter();
    }

    onCodeChange(codeChangedListener: (data: string) => void) {
        this._wordTextChangeEmitter.subscribe('change', codeChangedListener);
    }

    changeCode(code: string) {
        this._changingCode = true;
        this._wordTextChangeEmitter.suspendEventRecording();
        Word.run((ctx) => {
            ctx.document.body.insertText(code, 'Replace');
            return ctx.sync().then(() => {
                this._wordTextChangeEmitter.prevText = code;
                const html = ctx.document.body.getHtml();
                return ctx.sync().then(() => {
                    this._wordTextChangeEmitter.prevText = HtmlParser.parseHtml(html);
                    this._wordTextChangeEmitter.resumeEventRecording();
                    this._changingCode = false;
                });
            });
        }).catch(function () {
            this._changingCode = false;
            this._wordTextChangeEmitter.resumeEventRecording();
            console.error("Failed to check for change of text", arguments);
        }.bind(this));
    }
}

/**
 * Listens for the text change every intervalTime and if the text changes => emits 'change' event
 */
class WordTextChangeEmitter extends Emitter {
    public prevText: string;
    private _intervalTime: number;

    constructor(intervalTime = 500) {
        super()
        this._intervalTime = intervalTime;


        const changeChecker = () => {
            if (this._suspendEventRecording) {
                setTimeout(changeChecker, this._intervalTime);
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
                setTimeout(changeChecker, this._intervalTime);
                if (text !== this.prevText && !this._suspendEventRecording) {
                    console.log(text);
                    this.prevText = text;
                    this.emit('change', text);
                }
            }).catch(() => {
                setTimeout(changeChecker, this._intervalTime);
            });
        }
        const scheduleCheck = () => { setTimeout(changeChecker, this._intervalTime); };
        scheduleCheck();
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
 * Parser for ooxml word format to get text with line breaks. We need to use this, because context.document.body.text doesn't have any line breaks inside:(
 * @deprecated because ooxml retrieving is to slow(it's server based for Word online)
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