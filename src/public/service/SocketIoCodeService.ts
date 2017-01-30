import * as socketio from 'socket.io-client';

export interface CodeChangeSocketData {
    username: string,
    code: string
}

export class SocketIoCodeService {
    private static _instance: SocketIoCodeService = null;

    public static get instance(): SocketIoCodeService {
        if (!this._instance) {
            this._instance = new SocketIoCodeService();
        }
        return this._instance;
    }

    private _io: typeof socketio.Socket;

    constructor() {
        this._io = socketio.connect('https://sockscode.azurewebsites.net', { path: '/code' });
        this.onConnection(() => {
            console.log('SOCKET CONNECTED');
        }, () => {
            console.log('SOCKET DISCONECTED');
        })
    }

    changeCode(code: string) {
        this._io.emit('code change', code);
    }

    createRoom() {
        this._io.emit('create room');
    }

    joinRoom(roomUuid: string) {
        this._io.emit('join room', roomUuid);
    }

    onCodeChange(codeChangeFunc: (data: CodeChangeSocketData) => void) {
        this._io.on('code change', (data: CodeChangeSocketData) => {
            codeChangeFunc(data);
        });
    }

    onCreateRoom(roomCreatedFunc: (roomUuid: string) => void) {
        this._io.on('create room', (roomUuid: string) => {
            roomCreatedFunc(roomUuid);
        })
    }

    onConnection(onConnectionFunc: () => void, onDisconnectFunc: () => void) {
        this._io.on('connect', () => {
            onConnectionFunc();

        });
        this._io.on('disconnect', () => {
            onDisconnectFunc();
        });
    }
}