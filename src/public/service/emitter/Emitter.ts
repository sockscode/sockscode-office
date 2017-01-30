export type Listener = (data: any) => void;


export class Emitter {
    private _listeners: Map<String, Listener[]>;
    protected _suspendEventRecording: number;

    constructor() {
        this._listeners = new Map<String, Listener[]>();
        this._suspendEventRecording = 0;
    }

    suspendEventRecording() {
        this._suspendEventRecording++;
    }

    resumeEventRecording() {
        this._suspendEventRecording = Math.max(0, this._suspendEventRecording - 1);
    }

    isSuspendedEventRecording() {
        return !!this.suspendEventRecording;
    }

    emit(event: string, data: any) {
        let listenersList = this._listeners.get(event);
        listenersList && listenersList.forEach((listener) => {
            listener(data);
        })
    }

    subscribe(event: string, eventListener: Listener) {
        let listenersList = this._listeners.get(event);
        if (!listenersList) {
            listenersList = [] as Listener[];
            this._listeners.set(event, listenersList)
        }
        listenersList.push(eventListener);
    }

    unsubscribe(event: string, eventListener: Listener): boolean {
        let listenersList = this._listeners.get(event);
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