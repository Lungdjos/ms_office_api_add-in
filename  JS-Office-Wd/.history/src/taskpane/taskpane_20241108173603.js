/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    document.getElementById("insert-table").onclick = `insertTable`;
  }
});

// export async function run() {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */

//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";

//     await context.sync();
//   });
export async function insertTable() {
  return Word.run(async (context) => {
    const body = context.document.body;

    const table = body.tables.insertTable(5, 6, Word.InsertLocation.end);

    // set values for the table header
    table.getCell(0, 1).value = "Procurement Type";
    table.getCell(0, 2).value = "Procurement Method";
    table.getCell(0, 3).value = "Parts";
    table.getCell(0, 4).value = "Sections";
    table.getCell(0, 5).value = "Articles";
    table.getCell(0, 6).value = "Editable";

    // set values for the table body.

    await context.sync();
  });
}
