import * as React from 'react';
import CSSModules from 'react-css-modules';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { Button } from 'office-ui-fabric-react/lib/Button';

const styles = require("./Toolbar.css");

interface ToolbarProps {
    onCreateNewRoom: () => void;
    onRoomChange: (roomUuid: string) => void;
    onConnect: (roomUuid: string) => void;
    roomUuid: string;
}

interface ToolbarState {
    createNewRoomWindowOpen: boolean
}

@CSSModules(styles)
export class Toolbar extends React.Component<ToolbarProps, ToolbarState>{
    state = {
        createNewRoomWindowOpen: false
    }

    constructor() {
        super();
    }

    render() {
        const items = [{
            onClick: () => { },
            "key": "newItem", "name": "New", "icon": "Add", "ariaLabel": "New. Use left and right arrow keys to navigate", "data-automation-id": "newItemMenu", "subMenuProps": {
                "items": [
                    { onClick: () => { }, "key": "emailMessage", "name": "Email message", "icon": "Mail", "data-automation-id": "newEmailButton" },
                    { onClick: () => { }, "key": "calendarEvent", "name": "Calendar event", "icon": "Calendar" }
                ]
            }
        }, { "key": "upload", "name": "Upload", "icon": "Upload", "data-automation-id": "uploadButton", onClick: () => { console.log(arguments) } },
        { "key": "share", "name": "Share", "icon": "Share" },
        { "key": "download", "name": "Download", "icon": "Download" },
        { "key": "move", "name": "Move to...", "icon": "MoveToFolder" },
        { "key": "copy", "name": "Copy to...", "icon": "Copy" },
        { "key": "rename", "name": "Rename...", "icon": "Edit" },
        {
            key: 'upload1',
            name: 'Upload1',
            icon: 'Upload',
            onClick: () => alert('upload')
        },
        { "key": "disabled", "name": "Disabled...", "icon": "Cancel", "disabled": true }
        ];
        const farItems = [{ "key": "sort", "name": "Sort", "icon": "SortLines" }, { "key": "tile", "name": "Grid view", "icon": "Tiles" }, { "key": "info", "name": "Info", "icon": "Info" }];
        return <CommandBar isSearchBoxVisible={true}
            searchPlaceholderText='Room uuid' items={items} farItems={farItems}></CommandBar>
    }
}