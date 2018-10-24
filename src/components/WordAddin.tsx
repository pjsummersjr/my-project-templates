import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';

import * as OfficeHelpers from '@microsoft/office-js-helpers';
 
import { IOfficeAddinProps } from './OfficeAddin';

export interface IWordAddinProps extends IOfficeAddinProps {

}

interface IWordAddInState {
    data: any;
    mode: any;
}

export default class WordAddin extends React.Component<IWordAddinProps, IWordAddInState> {

    private provider: string = OfficeHelpers.DefaultEndpoints.AzureAD;;
    constructor(props: any) {
        super(props);
        this.state = {
            data: '',
            mode: 'LOADING_DATA'
        }
    }

    componentDidMount(): void {
        this.loadContent();
    }

    loadContent = () => {
        let self = this;

        let resourceUrl:string = "https://pjsummersjr2.ngrok.io/api/opportunities";
        self.props.authenticator.authenticate(this.provider, false).then(
            function(response: any){
                console.debug(`Requesting data from ${resourceUrl}`);
                let accessToken = response.access_token;
                var request = new XMLHttpRequest();
                request.addEventListener('load', () => {
                    console.log(`Response from server: ${JSON.stringify(this.responseText)}`);
                    self.setState({
                        data:this.responseText
                    })
                });
                request.open('GET', resourceUrl);
                request.setRequestHeader("Authorization", "Bearer " + accessToken);
                request.setRequestHeader("Content-Type", "application/json");
                request.setRequestHeader("Accept", "application/json");
                request.setRequestHeader("Cache-Control", "no-cache");
                request.send();
            },
            function(error: any) {
                console.log(`Error from authenticate: ${error}`);
                var token = self.props.authenticator.tokens.get(self.provider);
                console.log(`Got a token: ${JSON.stringify(token)}`);
            }
        )
        .catch(function(error:any){
            console.error(`Error caught in code: ${error}`);
        });
    }

    render() {
        if(!this.props.isOfficeInitialized) {
            return (<ProgressIndicator label="No Office environment detected" description="Please load this page within an Office add-in"/>);
        }
        return (<div>{JSON.stringify(this.state.data)}</div>);
    }

}