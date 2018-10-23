import * as React from 'react';
import * as ReactDOM from "react-dom";

import './index.scss';

let HelloWorld = () => {
  return (<h1>Hello there React World with Webpack!</h1>);
}

ReactDOM.render(
  <HelloWorld/>,
  document.getElementById("root")
);