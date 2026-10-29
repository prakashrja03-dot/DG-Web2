import React from "react";
import {
  useMsal,
  AuthenticatedTemplate,
  UnauthenticatedTemplate
} from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import EmployeeList from "./EmployeeList";
import "./App.css";

function App() {
  const { instance, accounts } = useMsal();

  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch((err) => console.error(err));
  };

  const handleLogout = () => {
    instance.logoutRedirect();
  };

  return (
    <div className="app-container">
      <UnauthenticatedTemplate>
        <div className="card">
          <h1>Dhash Safety App</h1>
          <p className="subtitle">Near Miss & Incident Reporting</p>
          <button className="btn btn-login" onClick={handleLogin}>
            Sign in with Microsoft
          </button>
        </div>
      </UnauthenticatedTemplate>

      <AuthenticatedTemplate>
        <div className="top-bar">
          <div>
            <h1 className="app-title">Dhash Safety App</h1>
            {accounts[0] && (
              <p className="user-meta">
                {accounts[0].name} &nbsp;|&nbsp; {accounts[0].username}
              </p>
            )}
          </div>
          <button className="btn btn-logout" onClick={handleLogout}>
            Sign out
          </button>
        </div>

        <EmployeeList />
      </AuthenticatedTemplate>
    </div>
  );
}

export default App;
