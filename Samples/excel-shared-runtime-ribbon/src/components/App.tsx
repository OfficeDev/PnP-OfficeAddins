import * as React from 'react';
import Header from './Header';
import ConnectButton from './ConnectButton';
import Progress from './Progress';
import OfficeAddinMessageBar from './OfficeAddinMessageBar';
import { getGlobal, ensureStateInitialized } from '../../utilities/office-apis-helpers';
import { btnConnectService } from '../commands/commands';
import DataFilter from './DataFilter';

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
    isStartOnDocOpen: boolean;
    isSignedIn: boolean;
}

export interface AppState {
    headerMessage?: string;
    errorMessage?: string;
}


export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);

        // Bind the methods that we want to pass to, and call in, a separate
        // module to this component. And rename setState to boundSetState
        // so code that passes boundSetState is more self-documenting.
        this.boundSetState = this.setState.bind(this);
       
        this.displayError = this.displayError.bind(this);
       
        
        console.log(btnConnectService);
       

      
            this.state = {
               
               
                headerMessage: 'Welcome',
                errorMessage: ''
            };
           
    
    }

    /*
        Properties
    */

    // The access token is not part of state because React is all about the
    // UI and the token is not used to affect the UI in any way.
    accessToken: string;

    /*
        Methods
    */

    boundSetState: () => {};

    setToken = (accesstoken: string) => {
        console.log('setting token');
        this.accessToken = accesstoken;
    }

    displayError = (error: string) => {
        this.setState({ errorMessage: error });
    }

    // Runs when the user clicks the X to close the message bar where
    // the error appears.
    errorDismissed = () => {
        this.setState({ errorMessage: '' });

       
    }

    dummy1 = async () => {
        
    }

    dummy2 = async () => {
        
    }

   

    render() {

        const { title, isOfficeInitialized } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/Onedrive_Charts_icon_80x80px.png'
                    message='Please sideload your add-in to see app body.'
                />
            );
        }

        // Set the body of the page based on where the user is in the workflow.
        let body;
        

        const g = getGlobal() as any;
        
        if (g.state.isConnected) {
            //connected UI
            // filter text button
            // preview data view
            // insert cf button
            body = (<DataFilter />);
        } else {
            //disconnected UI
            //just a connect button
            body = (<ConnectButton login={this.dummy2} />);
        }
       

        return (
            <div>
                {this.state.errorMessage ?
                    (<OfficeAddinMessageBar onDismiss={this.errorDismissed} message={this.state.errorMessage + ' '} />)
                    : null}

                <div className='ms-welcome'>
                    <Header logo='assets/Onedrive_Charts_icon_80x80px.png' title={this.props.title} message={this.state.headerMessage} />
                    {body}
                </div>
              
            </div>
        );
    }

    componentDidMount() {
        ensureStateInitialized();
        let g = getGlobal() as any;

        g.state.updateRct = (data: string) => {
            // `this` refers to our react component
            this.setState({ headerMessage: data });
        };
    }



}
