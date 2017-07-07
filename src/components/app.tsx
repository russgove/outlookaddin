import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { Header } from './header';
import { HeroList, HeroListItem } from './hero-list';
import * as pnp from "sp-pnp-js";
import * as adal from "adal-angular";
export interface AppProps {
    title: string;
}

export interface AppState {
    listItems: HeroListItem[];
}

export class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {

        super(props, context);
            this.state = {
            listItems: []
        };
        adal.
    }

    componentDidMount() {
        this.setState({
            listItems: [
                {
                    icon: 'Ribbon',
                    primaryText: 'Achieve more with Office integration'
                },
                {
                    icon: 'Unlock',
                    primaryText: 'Unlock features and functionality'
                },
                {
                    icon: 'Design',
                    primaryText: 'Create and visualize like a pro'
                }
            ]
        });
    }

    click = async () => {

        debugger;
              alert('hi')

    }

    render() {
        return (
           <iframe src="https://rgove3.sharepoint.com" />
        );
    };
};
