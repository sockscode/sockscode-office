import * as React from 'react';
import CSSModules from 'react-css-modules';
import autobind from 'autobind-decorator'

import { css } from 'office-ui-fabric-react/lib/Utilities';

const styles = require("./Info.css");

interface InfoProps {
}

interface InfoState {
}

@CSSModules(styles)
export class Info extends React.Component<InfoProps, InfoState>{
    constructor() {
        super();
    }

    render() {
        const infosNew = [
            'Click on the \'Connection\' button.',
            'Click on the \'Create new session\' button.',
            'Send generated room uuid to your teammate.'
        ];

        const infosExisting = [
            'Get generated room uuid to from your teammate. Paste this uuid to \'Room uuid\' input.',
            'Click on the \'Connection\' button.',
            'Click on the \'Connect to existing session\' button.',
        ];

        return <div className={styles.info}>
            <div className={css('ms-font-xxl')}>
                WELCOME
            </div>
            <div>
                This add-in enables connection with other sockscode enabled add-ins in order to have a pair programming session together with your teammate.
            </div>
            <div>
                To start a coding session follow this steps:
            </div>
            {infosNew.map((info, i) => {
                return this.renderInfo(i + 1, info);
            })}
            <div className={styles.or}>
                or
            </div>
            {infosExisting.map((info, i) => {
                return this.renderInfo(i + 1, info);
            })}
        </div>
    }

    renderInfo(index: number, text: string) {
        return <div className={styles.instruction}>
            <div className={styles.bullet}>
                {index}
            </div>
            <div className={styles['instruction-info']}>
                {text}
            </div>
        </div>
    }
}