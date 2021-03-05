// Copyright (c) Microsoft Corporation. All rights reserved. 
// Licensed under the MIT license. 

import React from 'react';
import './App.css';
import { login, tokenResponse, logout } from './auth/authUtil';
import { getPrinterShares } from "./graph/graphUtil";

const App: React.FunctionComponent<{}> = () => {

  const [authToken, setAuthToken] = React.useState<string>('');
  const [user, setUser] = React.useState<string>();
  const [message, setMessage] = React.useState<string>('');
  const [error, setError] = React.useState<string>();

  const onSignIn = async () => {
    try {
      const loginResp = await login();
      setAuthToken(loginResp.accessToken);
      setUser(loginResp.account.userName);
    } catch (error) {
      setError((error as Error).message);
    }
  };

  const onSignOut = () => {
    logout();
    setAuthToken('');
    setUser(undefined);
  }

  const onPrint = async () => {
    const response = await tokenResponse();
    setAuthToken(response.accessToken);
    const printShares = await getPrinterShares(authToken);
    let m:string = "";
    for( var i = 0; i < printShares.length; i++ ){

      m += printShares[i].id + ", " + printShares[i].name + "\n";

    }

    setMessage( m );
  };

  return (
    <section>
      <h1>
        Welcome to the Microsoft Universal Print Graph API demo
      </h1>
      {user === undefined ? (
        <button className="button" onClick={onSignIn}>Sign In</button>
      ) : (
          <>
            <button className="button" onClick={onSignOut}>
              Sign Out
            </button>
            <button className="button" onClick={onPrint}>
              プリンター一覧取得
            </button>
          </>
      )}
      {message && <p>{ message }</p>}
      {error && <p className="error">Error: {error}</p>}
    </section>
  );

};

export default App;
