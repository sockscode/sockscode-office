import { IOfficeService } from './OfficeService';
import { OfficeState } from '../reducers/Office'
import { Emitter } from './emitter/Emitter';
import { Parser } from 'xml2js';
import { stripPrefix } from 'xml2js/lib/processors';

interface CtxOutlineTuple extends Array<OneNote.RequestContext | OneNote.Outline> { 0: OneNote.RequestContext, 1: OneNote.Outline };

export class OneNoteService implements IOfficeService {
    private _outlineGetter: OutlineGetter;
    private _changeEmitter: OneNoteTextChangeEmitter;
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
        this._outlineGetter = new OutlineGetter();
        this._changeEmitter = new OneNoteTextChangeEmitter(this._outlineGetter);
    }

    onCodeChange(codeChangedListener: (data: string) => void) {
        this._changeEmitter.subscribe('change', codeChangedListener);
    }

    changeCode(code: string): void {
        this._changingCode = true;
        this._changeEmitter.suspendEventRecording();
        this._outlineGetter.getOutlineOrCreateOne().then((ctxAndOutlineTuple) => {
            const ctx = ctxAndOutlineTuple[0];
            const outline = ctxAndOutlineTuple[1];
            ctx.load(outline, 'paragraphs');
            return ctx.sync().then(() => {
                outline.paragraphs.items.forEach((p) => {
                    p.delete();
                });
                ctx.load(outline, 'paragraphs');
                return ctx.sync();
            }).then(() => {
                //we need this because OneNote doesn't support richText with '\n' proper way
                //@see http://stackoverflow.com/questions/41927348/adding-n-characters-inside-outline-paragraphs-items0-insertrichtextassibling
                const rows = code.split('\n');
                let richText = outline.paragraphs.items[0].insertRichTextAsSibling('Before', rows[0]);
                for (let i = 1; i < rows.length; i++) {
                    richText = richText.paragraph.insertRichTextAsSibling('After', rows[i]);
                }

                console.log('Set text to :' + code.trim());
                this._changeEmitter.prevText = code.trim(); //fixme
                return ctx.sync();
            });
        }).then(() => {
            this._changeEmitter.resumeEventRecording();
            this._changingCode = false;
        }).catch(() => {
            this._changeEmitter.resumeEventRecording();
            this._changingCode = false;
        });
    }
}

class OutlineGetter {
    private _outlineId: string = null;
    /**
     * Returns outline if it was created before and still exists or creates new outline.
     */
    public getOutlineOrCreateOne(): OfficeExtension.IPromise<CtxOutlineTuple> {
        const createNewOutline = (ctx: OneNote.RequestContext, activePage: OneNote.Page) => {
            //creating new outline
            const outline = activePage.addOutline(100, 100, '<p>//code goes here</p>');
            ctx.load(outline, 'id');
            return ctx.sync().then(() => {
                this._outlineId = outline.id;
                return [ctx, outline];
            });
        };

        return OneNote.run((ctx) => {
            const activePage = ctx.application.getActivePageOrNull();
            ctx.load(activePage, 'contents');
            return ctx.sync().then(() => {
                //fixme deal with noe activePage ?
                if (!this._outlineId) {
                    return createNewOutline(ctx, activePage);
                } else {
                    //trying to find old one, because it might have been removed
                    const pageContent = activePage.contents.items.find((pageContent: OneNote.PageContent) => {
                        return pageContent.id === this._outlineId && pageContent.type == 'Outline';
                    });
                    if (pageContent) {
                        return [ctx, pageContent.outline];
                    }
                    //no outline found :(. Let's craete new one.
                    return createNewOutline(ctx, activePage);
                }
            });
        });
    }
}

/**
 * Listens for the text change every intervalTime and if the text changes => emits 'change' event
 */
class OneNoteTextChangeEmitter extends Emitter {
    public prevText: string = '';
    private _intervalTime: number;
    private _outlineGetter: OutlineGetter;

    constructor(_outlineGetter: OutlineGetter, intervalTime = 500) {
        super();
        this._outlineGetter = _outlineGetter;
        this._intervalTime = intervalTime;

        const changeChecker = () => {
            //breaking early if suspended
            if (this._suspendEventRecording) {
                scheduleCheck();
                return;
            }
            let paragraphs: OneNote.Paragraph[];
            this._outlineGetter.getOutlineOrCreateOne().then((ctxAndOutlineTuple) => {
                //breaking early if suspended
                if (this._suspendEventRecording) {
                    scheduleCheck();
                    return;
                }
                const ctx = ctxAndOutlineTuple[0];
                const outline = ctxAndOutlineTuple[1];
                ctx.load(outline, 'paragraphs');
                return ctx.sync().then(() => {
                    paragraphs = outline.paragraphs.items.map((p) => {
                        p.load('id,type,richText/text');
                        return p;
                    });
                    return ctx.sync();
                });
            }).then(() => {
                //breaking early if suspended
                if (this._suspendEventRecording) {
                    scheduleCheck();
                    return;
                }
                const text = paragraphs.filter((p) => {
                    return p.type === 'RichText';
                }).map((p) => {
                    return p.richText.text;
                }).join('\n');
                if (this.prevText.trim() !== text.trim()) { //fixme 
                    this.prevText = text;
                    console.log('FOUND CHANGE from "' + this.prevText.trim() + '" to "' + this.prevText.trim() + '"');
                    this.emit('change', text);
                }
                scheduleCheck();
            }).catch((error) => {
                console.log('Failed to check for change', error);
                scheduleCheck();
            });
        }

        const scheduleCheck = () => { setTimeout(changeChecker, this._intervalTime); };
        scheduleCheck();
    }
}