import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { Header } from './header';
import { HeroList, HeroListItem } from './hero-list';
import * as pnp from 'sp-pnp-js';
import * as adal from 'adal-angular';
export interface AppProps {
    title: string;
}
import adalConfig from '../AdalConfig';
import { IAdalConfig } from '../../IAdalConfig';
export interface AppState {
    listItems: HeroListItem[];
    loading: boolean;
    error: string;
    signedIn: boolean;
}

export class App extends React.Component<AppProps, AppState> {
    debugger;
    private authCtx: adal.AuthenticationContext;
    constructor(props, context) {
        super(props, context);
        debugger;
        this.state = {
            listItems: [],
            signedIn: false,
            error: '',
            loading: false,
        };
        const config: IAdalConfig = adalConfig;
        this.authCtx = new adal(config);
        window['Logging'] = {
            level: 10,
            log: function (message) {
                console.log('ADAL MESSAGGE ' + message);
            }
        };
        
        if (!this.authCtx.isCallback(window.location.hash)) {
            this.authCtx.login();
        } else {
            console.log('user is ' + this.authCtx.getCachedUser().userName);
            this.authCtx.handleWindowCallback();
            this.authCtx.acquireToken('f8f8d2ad-7c9d-4aac-80eb-3f00a263c879', (error, token) => {
                if (error) {
                    console.log(error);
                    return;
                }
                let myHeaders = new Headers();
                myHeaders.append('Authorization', 'Bearer ' + token);
                let myInit = {
                    method: 'GET',
                    headers: myHeaders,
                    mode: 'cors' as RequestMode,
                    cache: 'default' as RequestCache
                };
                fetch("https://rgove3.sharepoint.com/cboxTest/_api/web/lists/getByTitle('ideas')/items", myInit).then((response) => {
                    debugger;
                    console.log('GOT WEBw ' + response);
                }).catch((error) => {
                    console.log('Error: ' + error);
                    debugger;
                });
            });
        }
    }
    
    render() {
        debugger;
        if (this.authCtx.loginInProgress()) {
            return <div>logging in , please wait.</div>
        }
        return (
            <div>logged in as {this.authCtx.getCachedUser().userName}</div>
        );
    };
};
