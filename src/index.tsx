import * as React from 'react';
import * as ReactDOM from "react-dom";
import SimpleService from './services/SimpleService';

import './index.scss';

let HelloWorld = () => {
  let service: SimpleService = new SimpleService();
  let message: string = service.getSimpleData()[0];
  return (<h1>Hello there {message}! Welcome to the React World with Webpack!</h1>);
}

ReactDOM.render(
  <HelloWorld/>,
  document.getElementById("root")
);