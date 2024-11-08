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

    // Dummy data for the table (for example purposes)
    const rows = [
      ["Type A", "Method X", "Part 1", "Section 1", "Article 1", "true"],
      ["Type B", "Method Y", "Part 2", "Section 2", "Article 2", "false"],
      ["Type C", "Method Z", "Part 3", "Section 3", "Article 3", "true"],
      ["Type D", "Method W", "Part 4", "Section 4", "Article 4", "false"],
      ["Type E", "Method V", "Part 5", "Section 5", "Article 5", "true"],
    ];

    // Insert data into table and set editable flag
    for (let i = 0; i < rows.length; i++) {
      for (let j = 0; j < rows[i].length; j++) {
        table.getCell(i + 1, j).value = rows[i][j]; // +1 because row 0 is the header
      }

      // Check the 'Editable' flag in the last column and set the row as editable or read-only
      const editableFlag = rows[i][5] === "true";
        // Make the row read-only by disabling editing (you can also apply other styles here)
        for (let j = 0; j < table.rowCount; j++) {
          table.getRow(i + 1).getCell(j).contentControlProperties.locked = true;
        }
      }
    }

    await context.sync();
  });
}
