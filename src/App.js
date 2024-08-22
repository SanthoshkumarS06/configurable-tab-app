import { useEffect } from "react";
import logo from "./logo.svg";
import "./App.css";

import * as microsoftTeams from "@microsoft/teams-js";

function App() {
  useEffect(() => {
    microsoftTeams.app
      .initialize()
      .then(() => {
        microsoftTeams.app.getContext().then((context) => {
          console.log("context", context);
        });
      })
      .catch(console.error);
  }, []);

  const handleClick = () => {
    microsoftTeams.pages.config.registerOnSaveHandler((saveEvent) => {
      const configPromise = microsoftTeams.pages.config.setConfig({
        suggestedDisplayName: "CrudTeamsWebpart2",
        entityId: "a371a734-4260-4481-9b4d-73b90ffbe577",
        contentUrl:
          "https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId=a371a734-4260-4481-9b4d-73b90ffbe577%26forceLocale={locale}",
        websiteUrl:
          "https://products.office.com/en-us/sharepoint/collaboration",
      });
      configPromise
        .then((result) => {
          saveEvent.notifySuccess();
        })
        .catch((error) => {
          saveEvent.notifyFailure("failure message");
        });
    });
  };

  return (
    <div className="App">
      <button onClick={handleClick}>Register On Save Handler</button>
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Edit <code>src/App.js</code> and save to reload.
        </p>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React
        </a>
      </header>
    </div>
  );
}

export default App;
