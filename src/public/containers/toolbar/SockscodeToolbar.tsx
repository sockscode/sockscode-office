import { connect } from 'react-redux';
import { Dispatch } from 'redux';
import { SockscodeState } from '../../reducers/Reducers';
import { codeChanged } from '../../actions/Actions';
import { Toolbar } from '../../components/toolbar/Toolbar';
import { createRoom, changedRoom, joinRoom, createTextControll, disconnect } from '../../actions/Actions';

const mapStateToProps = (state: SockscodeState) => {
    const { roomUuid } = state;
    return { roomUuid };
}

const mapDispatchToProps = (dispatch: Dispatch<{}>) => {
    return {
        onCreateNewRoom: () => {
            dispatch(createRoom());
        },
        onRoomChange: (roomUuid: string) => {
            dispatch(changedRoom(roomUuid));
        },
        onConnect: (roomUuid: string) => {
            dispatch(joinRoom(roomUuid));
        },
        onCreateTextControll: () => {
            dispatch(createTextControll());
        },
        onDisconnect: () => {
            dispatch(disconnect());
        }
    }
}
const SockscodeToolbar = connect(
    mapStateToProps,
    mapDispatchToProps
)(Toolbar);

export { SockscodeToolbar }