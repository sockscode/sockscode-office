import * as React from 'react'
import CSSModules from 'react-css-modules';
import { SockscodeToolbar } from './toolbar/SockscodeToolbar';
import { createStore } from 'redux';
import { Provider } from 'react-redux'
import { sockscodeApp } from '../reducers/Reducers';
import { SocketIoCodeService } from '../service/SocketIoCodeService';
import { createdRoom, codeChanged, remoteCodeChanged } from '../actions/Actions'

interface AppProps {

}

interface AppState {

}

const store = createStore(sockscodeApp);
const styles = require("./App.css");
const socketIoCodeService = SocketIoCodeService.instance;
socketIoCodeService.onCreateRoom((roomUuid) => {
    store.dispatch(createdRoom(roomUuid));
});
socketIoCodeService.onCodeChange((codeChangeSocketData) => {
    store.dispatch(remoteCodeChanged(codeChangeSocketData.code));
});

@CSSModules(styles)
export class App extends React.Component<AppProps, AppState>{
    constructor() {
        super();
    }

    render() {
        return <Provider store={store}>
            <div className={styles.container}>
                <div>
                    <SockscodeToolbar />
                </div>
            </div>
        </Provider>
    }
}