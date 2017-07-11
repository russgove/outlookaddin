import * as React from 'react';
import { Button, ButtonType, Label, DetailsList, PrimaryButton } from 'office-ui-fabric-react';
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
    body: string;
    from: string;
    attachments: Array<any>
}

export class App extends React.Component<AppProps, AppState> {
    debugger;
    private serviceRequest: any;
    private authCtx: adal.AuthenticationContext;
    constructor(props, context) {
        super(props, context);
        debugger;
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
            this.authCtx.handleWindowCallback();
            console.log('user is ' + this.authCtx.getCachedUser().userName);
            debugger;
            this.state = {
                listItems: [],
                signedIn: false,
                error: '',
                loading: false,
                body: '',
                attachments: Office.context.mailbox.item.attachments,
                from: Office.context.mailbox.item.from.emailAddress,
            };
            Office.context.mailbox.item.body.getAsync(
                "html",
                { asyncContext: "This is passed to the callback" },
                (result) => {
                    debugger;
                    this.setState({ body: result.value });
                }

            );

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
    private renderBody() {
        return { __html: this.state.body };
    }
    public editDocument(item: any): void {
        debugger;
        this.authCtx.acquireToken('f8f8d2ad-7c9d-4aac-80eb-3f00a263c879', (error, token) => {
            debugger;
            if (error) {
                console.log(error);
                return;
            }

            let myHeaders = new Headers();
            debugger;
            myHeaders.append('Authorization', 'Bearer ' + token);
            let settings = {
                method: 'GET',
                headers: myHeaders,
                mode: 'cors' as RequestMode,
                cache: 'no-cache' as RequestCache
            };
            var url = Office.context.mailbox.restUrl + '/v2.0/me/messages';
            fetch(url, settings).then((response) => {
                debugger;
            }).catch((err) => {
                debugger;
            })

        });

    }
    private getAttachment(asyncResult) {
        if (asyncResult.status === "succeeded") {
            let myHeaders = new Headers();
            debugger;
            myHeaders.append('Authorization', 'Bearer ' + asyncResult.value);
            let settings = {
                method: 'GET',
                headers: myHeaders,
                mode: 'cors' as RequestMode,
                cache: 'no-cache' as RequestCache
            };
            var url = Office.context.mailbox.restUrl + '/v2.0/me/MailboxSettings/AutomaticRepliesSetting';
            fetch(url, settings).then((response) => {
                debugger;
            }).catch((err) => {
                debugger;
            })


            //   makeServiceRequest();


        }
    }
    render() {
        debugger;
        if (this.authCtx.loginInProgress()) {
            return <div>logging in , please wait.</div>
        }
        return (
            <div>
                <Label>Login info</Label>
                <div>logged in as {this.authCtx.getCachedUser().userName}</div>
                <Label>Message Text</Label>
                <div dangerouslySetInnerHTML={this.renderBody()} />
                <Label>Message From</Label>
                <div>{this.state.from}</div>
                <Label>Attachments</Label>
                <DetailsList items={this.state.attachments}
                    columns={[
                        {
                            key: "Edit", name: "", fieldName: "Title", minWidth: 20,
                            onRender: (item) => <div>
                                <i onClick={(e) => { this.editDocument(item); }}
                                    className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
                            </div>
                        },
                        { key: "name", name: "name", fieldName: "name", minWidth: 20, maxWidth: 70 },
                        { key: "size", name: "size", fieldName: "size", minWidth: 20, maxWidth: 70 },
                        { key: "isInline", name: "isInline", fieldName: "isInline", minWidth: 20, maxWidth: 70 },
                        { key: "attachmentType", name: "attachmentType", fieldName: "attachmentType", minWidth: 20, maxWidth: 70 },
                        { key: "id", name: "id", fieldName: "id", minWidth: 20, maxWidth: 70 },

                    ]}
                />
                <PrimaryButton href="#" icon="ms-Icon--Save">
                    <i className="ms-Icon ms-Icon--Save" aria-hidden="true"></i>
                    Save
        </PrimaryButton>
            </div>
        );
    };

};
