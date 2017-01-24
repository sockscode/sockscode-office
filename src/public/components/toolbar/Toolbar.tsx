import * as React from 'react';
import CSSModules from 'react-css-modules';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { Button } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IconName } from 'office-ui-fabric-react/lib/Icon';
import { ContextualMenu, DirectionalHint } from 'office-ui-fabric-react/lib/ContextualMenu';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import autobind from 'autobind-decorator'

import { css } from 'office-ui-fabric-react/lib/Utilities';

const styles = require("./Toolbar.css");

interface ToolbarProps {
    onCreateNewRoom: () => void;
    onRoomChange: (roomUuid: string) => void;
    onConnect: (roomUuid: string) => void;
    roomUuid: string;
}

interface ToolbarState {
    createNewRoomWindowOpen: boolean,
    contextMenuTarget: any,
    isContextMenuVisible: boolean
}

@CSSModules(styles)
export class Toolbar extends React.Component<ToolbarProps, ToolbarState>{
    state = {
        createNewRoomWindowOpen: false,
        contextMenuTarget: null as any,
        isContextMenuVisible: false
    }

    constructor() {
        super(); console.log(CommandBar);
    }

    render() {
        return <div className={css('ms-CommandBar', )} ref='commandBarRegion'>
            <FocusZone className={styles.toolbar} ref='focusZone' direction={FocusZoneDirection.horizontal} rootProps={{ role: 'menubar' }}>
                {this._renderLogo()}
                <TextField className={styles.roomInput}
                    value={this.props.roomUuid}
                    onChanged={this.props.onRoomChange}
                    placeholder='Room uuid'
                    />
                <div>
                    <Button onClick={this._onConnectButtonClick} id='connect'> Connection</Button>
                    {this.state.isContextMenuVisible ? (
                        <ContextualMenu
                            shouldFocusOnMount={true}
                            target={this.state.contextMenuTarget}
                            onDismiss={this._onDismiss}
                            directionalHint={DirectionalHint.bottomLeftEdge}
                            items={
                                [
                                    {
                                        name: 'New Connection',
                                        key: 'newItem',
                                        iconProps: {
                                            iconName: 'Add' as IconName
                                        },
                                        subMenuProps: {
                                            items: [
                                                {
                                                    key: 'newSession',
                                                    name: 'Create new session.',
                                                    title: 'Create new session.',
                                                    onClick: this._onCreateNewRoom
                                                },
                                                {
                                                    key: 'existing',
                                                    name: 'Connect to existing session',
                                                    title: 'Connect to existing session. You\'ll need roomUuid for this.',
                                                    onClick: this._onConnectToExisting
                                                }
                                            ],
                                        }
                                    }, {
                                        name: 'Stop Connection',
                                        key: 'stopConnection',
                                        iconProps: {
                                            iconName: 'Cancel' as IconName
                                        },
                                    }
                                ]
                            }
                            />) : (null)
                    }
                </div>
            </FocusZone>
        </div>
    }
    
    @autobind
    private _onCreateNewRoom() {
        this.props.onCreateNewRoom();
    }
    
    @autobind
    private _onConnectToExisting() {
        this.props.onConnect(this.props.roomUuid);
    }

    @autobind
    private _onConnectButtonClick(event: React.MouseEvent<any>) {
        this.setState({ contextMenuTarget: event.currentTarget, isContextMenuVisible: true } as any);
    }

    @autobind
    private _onDismiss(event: any) {
        this.setState({ isContextMenuVisible: false } as any);
    }

    private _renderLogo() {
        return <svg className={styles.logo} viewBox="0 0 200 200" width="30" height="30" xmlns="http://www.w3.org/2000/svg" version="1.1">
            <defs>
                <pattern id="sock" x="0" y="0" width="90" height="30" patternUnits="userSpaceOnUse">
                    <rect width="90" height="30" fill="#F35325"></rect>
                    <rect width="90" height="18" fill="#81BC06"></rect>
                </pattern>
                <pattern id="sockRotated" x="0" y="0" width="90" height="30" patternUnits="userSpaceOnUse">
                    <rect width="90" height="30" fill="#05A6F0"></rect>
                    <rect width="90" height="18" fill="#FFBA08"></rect>
                </pattern>
            </defs>
            <circle cx="100" cy="100" r="98" stroke="black" fill="white" strokeWidth="0"></circle>
            <g transform="translate(44,23)">
                <g>
                    <path d="M 0 0 L 46 0 L 46 85 L 85.23138726188375 102.50323431684045 A 23 23 0 0 1 94.54367554580823 136.9213090611957 L 94.15441245192645 137.4634969419596 A 23 23 0 0 1 65.6897146442473 144.8662514593451 L 11.494563534719532 119.40105997414531 A 20 20 0 0 1 0 101.29974647111253 Z" style={{ stroke: '#000000', fill: 'url(#sock)' }} strokeWidth="0"></path>
                </g>
                <g transform="rotate(180,56,77)">
                    <path d="M 0 0 L 46 0 L 46 85 L 85.23138726188375 102.50323431684045 A 23 23 0 0 1 94.54367554580823 136.9213090611957 L 94.15441245192645 137.4634969419596 A 23 23 0 0 1 65.6897146442473 144.8662514593451 L 11.494563534719532 119.40105997414531 A 20 20 0 0 1 0 101.29974647111253 Z" style={{ stroke: '#000000', fill: 'url(#sockRotated)' }} strokeWidth="0"></path>
                </g>
            </g>
        </svg>
    }
}