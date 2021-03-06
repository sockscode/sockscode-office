import { Dispatch } from 'redux';
import { SocketIoCodeService } from '../service/SocketIoCodeService';
import { OfficeState } from '../reducers/Office';
//import { OfficeService } from '../service/OfficeService';
/*
 * action types
 */
export const CREATE_ROOM = 'CREATE_ROOM';
export const CREATED_ROOM = 'CREATED_ROOM';
export const CHANGED_ROOM = 'CHANGED_ROOM';
export const JOIN_ROOM = 'JOIN_ROOM';
export const DISCONNECT = 'DISCONNECT';
export const CODE_CHANGED_REMOTE = 'CODE_CHANGED_REMOTE';
export const CODE_CHANGED_LOCAL = 'CODE_CHANGED_LOCAL';
/**
 * Office action types
 */
export const OFFICE_INITIALIZED = 'OFFICE_INITIALIZED';
export const CREATE_TEXT_CONTROLL = 'CREATE_TEXT_CONTROLL';


interface HasRoomUuid {
    roomUuid: string;
}

interface HasCode {
    code: string;
}

export interface CreateRoomAction {
    type: 'CREATE_ROOM';
}
export interface ChangedRoomAction extends HasRoomUuid {
    type: 'CHANGED_ROOM';
}
export interface CreatedRoomAction extends HasRoomUuid {
    type: 'CREATED_ROOM';
}
export interface JoinRoomAction extends HasRoomUuid {
    type: 'JOIN_ROOM';
}

export interface CodeChangedLocalAction extends HasCode {
    type: 'CODE_CHANGED_LOCAL';
}
export interface CodeChangedRemoteAction extends HasCode {
    type: 'CODE_CHANGED_REMOTE';
}

export interface DisconnectAction {
    type: 'DISCONNECT';
}

/**
 * Office action interfaces
 */
export interface OfficeInitStateAction extends OfficeState {
    type: 'OFFICE_INITIALIZED',
}

export interface CreateTextControll {
    type: 'CREATE_TEXT_CONTROLL'
}

/*
 * action creators
 */
export function createRoom(): CreateRoomAction {
    SocketIoCodeService.instance.createRoom();
    return { type: CREATE_ROOM }
}

export function joinRoom(roomUuid: string): JoinRoomAction {
    SocketIoCodeService.instance.joinRoom(roomUuid);
    return { type: JOIN_ROOM, roomUuid };
}

export function createdRoom(roomUuid: string): CreatedRoomAction {
    return { type: CREATED_ROOM, roomUuid }
}

export function changedRoom(roomUuid: string): ChangedRoomAction {
    return { type: CHANGED_ROOM, roomUuid };
}

export function disconnect(): DisconnectAction {
    SocketIoCodeService.instance.close();
    return { type: DISCONNECT };
}

export function codeChanged(code: string): CodeChangedLocalAction {
    SocketIoCodeService.instance.changeCode(code);
    return { type: CODE_CHANGED_LOCAL, code };
}

export function remoteCodeChanged(code: string): CodeChangedRemoteAction {
    return { type: CODE_CHANGED_REMOTE, code };
}

/**
 * Office action creators
 */
export function officeInitialized(officeState: OfficeState): OfficeInitStateAction {
    return { ...officeState, type: 'OFFICE_INITIALIZED' };
}

export function createTextControll() {
    // OfficeService.addContentControls().catch((error) => {
    //     console.error(error);
    // });
    return { type: CREATE_TEXT_CONTROLL };
}

