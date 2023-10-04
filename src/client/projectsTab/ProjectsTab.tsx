import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app } from "@microsoft/teams-js";

/**
 * Implementation of the Projects Tab content page
 */
export const ProjectsTab = () => {

  const [{ inTeams, theme, context }] = useTeams();
  const [entityId, setEntityId] = useState<string | undefined>();

  const [messages, setMessages] = useState<string[]>([]);

  useEffect(() => {
    if (inTeams === true) {

      window.addEventListener('message', event => {
        setMessages(messages => [JSON.stringify(event.data), ...messages]);
      }, false);

      app.notifySuccess();
    } else {
      setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams]);

  useEffect(() => {
    if (context) {
      setEntityId(context.page.id);
    }
  }, [context]);

  /**
   * The render() method to create the UI of the tab
   */
  return (
    <Provider theme={theme}>
      <iframe style={{ width: '100%', height: '50vh' }} src="https://hubblecontent.osi.office.net/contentsvc/m365contentpicker/index.html?p=3&app=1001&aud=prod&channel=devmain&setlang=${language}&msel=0&env=prod&premium=1${themesColor}" />
      <div style={{ width: '100%', height: '20vh', overflow: 'auto' }}>
        {messages.map(m => <div>{m}</div>)}
                        </div>
    </Provider>
  );
};
