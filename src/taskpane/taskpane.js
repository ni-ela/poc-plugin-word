/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import React from 'react';
import ReactDOM from 'react-dom';
import MyComponent from './components/Button';
import { Button, FluentProvider, teamsLightTheme } from '@fluentui/react-components';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Certifique-se de que o Office.js está pronto antes de renderizar o aplicativo React
    const App = () => (
      <FluentProvider theme={teamsLightTheme}>
        <Button />
      </FluentProvider>
    );

    ReactDOM.render(<App />, document.getElementById('root'));

    //  document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
