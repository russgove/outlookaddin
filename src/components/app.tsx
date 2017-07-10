import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import { Header } from "./header";
import { HeroList, HeroListItem } from "./hero-list";
import * as pnp from "sp-pnp-js";
import * as adal from "adal-angular";
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
    oauth_id_token: string;
    oauth_state: string;
    oauth_session_state: string;

}

export class App extends React.Component<AppProps, AppState> {
    debugger;



    private authCtx: adal.AuthenticationContext;
    private getQueryVariable(variable: string) {
        let hash = window.location.hash.substring(1);
        let vars = hash.split('&');
        for (var i = 0; i < vars.length; i++) {
            var pair = vars[i].split('=');
            if (decodeURIComponent(pair[0]) == variable) {
                return decodeURIComponent(pair[1]);
            }
        }
        console.log('Query variable %s not found', variable);
        return null;
    }
    constructor(props, context) {

        super(props, context);
        debugger;
        this.state = {
            listItems: [],
            signedIn: false,
            error: "",
            loading: false,
            oauth_id_token: this.getQueryVariable("id_token"),
            oauth_session_state: this.getQueryVariable("session_state"),
            oauth_state: this.getQueryVariable("state"),

        };
        const config: IAdalConfig = adalConfig;

        //config.webPartId = this.props;
        config.callback = (error: any, token: string): void => {
            debugger;
            this.setState((previousState: AppState, currentProps: AppProps): AppState => {
                previousState.error = error;
                previousState.signedIn = !(!this.authCtx.getCachedUser());
                return previousState;
            });
        };
        debugger;
        this.authCtx = new adal(config);
        // this.authCtx.prototype._singletonInstance = undefined;
        if (!this.state.oauth_id_token) {
            this.authCtx.login();
        } else {
            console.log("user is " + this.authCtx.getCachedUser());
              this.authCtx.handleWindowCallback();
            const token=this.authCtx.acquireToken('f8f8d2ad-7c9d-4aac-80eb-3f00a263c879');
            pnp.setup({
                headers: {
                    'Authorization':'Bearer '+ token,
                },
                baseUrl: "https://rgove3.sharepoint.com"

            });
            debugger;
            pnp.sp.web.lists.getByTitle('Contacts').items.get().then((items)=>{
                debugger;
                console.log("GOT WEB "+items[0]["Title"]);
            }).catch((err)=>{
                debugger;
                console.log("Error "+err)
            });
        }



    }

    componentDidMount() {
        debugger;
        this.authCtx.handleWindowCallback();

        if (window !== window.top) {
            return;
        }

        this.setState((previousState: AppState, props: AppProps): AppState => {
            previousState.error = this.authCtx.getLoginError();
            previousState.signedIn = !(!this.authCtx.getCachedUser());
            return previousState;
        });

    }

    click = async () => {

        debugger;
        alert("hi")

    }

    render() {
        if (this.state.oauth_id_token === null) {
            return <div>logging in , please wait.</div>
        }
        return (
            <div>logged in as {this.authCtx.getCachedUser().userName}</div>
        );
    };
};
