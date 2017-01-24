import * as React from 'react'
import CSSModules from 'react-css-modules';
import { SockscodeToolbar } from './toolbar/SockscodeToolbar';
import { createStore } from 'redux';
import { Provider } from 'react-redux'
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { sockscodeApp } from '../reducers/Reducers';
import { SocketIoCodeService } from '../service/SocketIoCodeService';
import { OfficeService } from '../service/OfficeService';
import { createdRoom, codeChanged, remoteCodeChanged } from '../actions/Actions'
import { css } from 'office-ui-fabric-react/lib/Utilities';

interface AppProps {

}

interface AppState {
    officeInitialized: boolean
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
const officeService: OfficeService = new OfficeService();

@CSSModules(styles)
export class App extends React.Component<AppProps, AppState>{

    state = {
        officeInitialized: false
    }
    constructor() {
        super();
        officeService.promiseInitialize().then(() => {
            this.setState({ officeInitialized: true });
        });
        //fixme catch? fixme redux
    }

    render() {
        return <Provider store={store}>
            <div className={css(styles.container, 'ms-Fabric', 'ms-font-m')}>
                {this.state.officeInitialized ? <div>
                    <SockscodeToolbar />
                </div> : <Spinner className={styles.spinner} type={SpinnerType.large} label='Initializing...' />}
            </div>
        </Provider>
    }
}