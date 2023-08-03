/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

function getAutoShowElement() {
  return document.getElementById("autoshow") as HTMLInputElement
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    const autoShowElement = getAutoShowElement();
    autoShowElement.onclick = setAutoShow;
    getAutoShow().then(function (state) {
      autoShowElement.checked = state;
    })
  }
});

export async function getAutoShow() {
  return new Promise<boolean>((resolve) => {
    const getAutoShowTaskpaneWithDocument = (context: Office.Context) => {
      const value = context.document.settings.get("Office.AutoShowTaskpaneWithDocument");
      const state = (typeof value === "boolean") ? (value as boolean) : false;
      console.log(`auto show is: ${state}`);
      resolve(state);
    }
    getAutoShowTaskpaneWithDocument(Office.context);
  })
}

export async function setAutoShow() {
  return new Promise<void>((resolve) => {
    const setAutoShowTaskpaneWithDocument = (context: Office.Context) => {
      const value = getAutoShowElement().checked;
      context.document.settings.set("Office.AutoShowTaskpaneWithDocument", value);
      context.document.settings.saveAsync(function () {
        console.log(`set auto show to: ${value}`);
        resolve();
      })
    }
    setAutoShowTaskpaneWithDocument(Office.context);
  })
}

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";
    paragraph.font.size = 24;

    await context.sync();
  });
}
