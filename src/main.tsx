import * as React from 'react';
import { render } from 'react-dom';
import { App } from './components/app';
import { Progress } from './components/progress';
import './assets/styles/global.scss';

(() => {
    const title = 'Git the Gist';
    const container = document.querySelector('#container');
    debugger;
    /* Render application after Office initializes */
    Office.initialize = () => {
    //    debugger;
        render(
            <App title={title} />,
            container
        );
    };

    /* Initial render showing a progress bar */
    render(<Progress title={title} logo='assets/logo-filled.png' message='Please sideload your addin to see app body.' />, container);
})();

