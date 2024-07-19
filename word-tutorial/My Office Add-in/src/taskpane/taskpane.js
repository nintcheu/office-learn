/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { base64Image } from "../../base64Image";
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    document.getElementById("validate-paragraph").onclick = () => tryCatch(checkSensitviteInformation);

    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    document.getElementById("apply-custom-style").onclick = () => tryCatch(applyCustomStyle);
    document.getElementById("change-font").onclick = () => tryCatch(changeFont);
    document.getElementById("insert-text-into-range").onclick = () => tryCatch(insertTextIntoRange);
    document.getElementById("insert-text-outside-range").onclick = () => tryCatch(insertTextBeforeRange);
    document.getElementById("replace-text").onclick = () => tryCatch(replaceText);
    document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
    document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
    document.getElementById("insert-table").onclick = () => tryCatch(insertTable);
    document.getElementById("create-content-control").onclick = () => tryCatch(createContentControl);
    document.getElementById("replace-content-in-control").onclick = () => tryCatch(replaceContentInControl);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function insertParagraph() {
  await Word.run(async (context) => {

    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
                            Word.InsertLocation.start);

      await context.sync();
  });

}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
  }
}

async function applyStyle() {
  await Word.run(async (context) => {

    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;

      await context.sync();
  });
}

async function applyCustomStyle() {
  await Word.run(async (context) => {

    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";

      await context.sync();
  });
}

async function changeFont() {
  await Word.run(async (context) => {

    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });

      await context.sync();
  });
}

async function insertTextIntoRange() {
  await Word.run(async (context) => {

    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (M365)", Word.InsertLocation.end);

    originalRange.load("text");
    await context.sync();

      doc.body.insertParagraph("Original range: " + originalRange.text, Word.InsertLocation.end);

      await context.sync();
  });
}

async function insertTextBeforeRange() {
  await Word.run(async (context) => {

    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", Word.InsertLocation.before);

    originalRange.load("text");
    await context.sync();
    
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, Word.InsertLocation.end);
    
    await context.sync();

  });
}

async function replaceText() {
  await Word.run(async (context) => {

    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", Word.InsertLocation.replace);

      await context.sync();
  });
}

async function insertImage() {
  await Word.run(async (context) => {

    context.document.body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);

      await context.sync();
  });
}

async function insertHTML() {
  await Word.run(async (context) => {

    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', Word.InsertLocation.end);

      await context.sync();
  });
}

async function insertTable() {
  await Word.run(async (context) => {

    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();

    const tableData = [
      ["Name", "ID", "Birth City"],
      ["Bob", "434", "Chicago"],
      ["Sue", "719", "Havana"],
  ];
secondParagraph.insertTable(3, 3, Word.InsertLocation.after, tableData);

      await context.sync();
  });
}

async function createContentControl() {
  await Word.run(async (context) => {

    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";

      await context.sync();
  });
}

async function replaceContentInControl() {
  await Word.run(async (context) => {

    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", Word.InsertLocation.replace);

      await context.sync();
  });
}

      //When verify-content-button is selected, get the content of the email body
      async function checkSensitviteInformation () {
        //addInformational("Any sensitive data found in your mail. Great job! You respect our policies.");
        //showError("You found sensitive date in your mail. You don't respect the policies. Please, take action to correct the issue.");

      // Gets the text content of the body.
      // Run a batch operation against the Word object model.
        await Word.run(async (context) => {

            //if (result.status == "succeeded") {

              const patterns = {
                ssn: /\b(?!000|666|9\d{2})([0-8]\d{2}|7([0-6]\d|7[012]))([-]?)\d{2}\3\d{4}\b/,
                creditCard: /\b(?:4[0-9]{12}(?:[0-9]{3})?|5[1-5][0-9]{14}|3[47][0-9]{13}|3(?:0[0-5]|[68][0-9])[0-9]{11}|6(?:011|5[0-9]{2})[0-9]{12}|(?:2131|1800|35\d{3})\d{11})\b|\b(?:(?:4[0-9]{3}|5[1-5][0-9]{2}|6[0-9]{3}|3[47][0-9]{2})[- ]?[0-9]{4}[- ]?[0-9]{4}[- ]?[0-9]{4})\b/,
                dateOfBirth: /\b(0[1-9]|1[0-2])[/-](0[1-9]|[12]\d|3[01])[/-](19|20)\d{2}\b/,
                email: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/,
                phoneNumber: /\b(\+\d{1,2}\s?)?1?\-?\.?\s?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}\b/,
                ipAddress: /\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b/,
                passportNumber: /\b[A-Z]{1,2}[0-9]{6,9}\b/,
                driverLicense: /\b[A-Z]{1,2}[0-9]{5,7}\b/,
                bankAccount: /\b[0-9]{8,17}\b/
              };

              const sensitiveInfo = {};
              // Create a proxy object for the document body.
              const body = context.document.body;

              let text = body.text;

              for (const [type, pattern] of Object.entries(patterns)) {
                const matches = text.match(pattern);
                if (matches) {
                  sensitiveInfo[type] = matches.map(match => ({
                    value: match,
                    index: text.indexOf(match)
                  }));
                  console.log(type.toUpperCase(), " sensitive");
                  console.log(type.toUpperCase() + ": " + sensitiveInfo[type][0].value);

                  //body.load(type.toUpperCase() + " detected as sensitive data");
                  //addInformational(type.toUpperCase() + " detected as sensitive data");
                  //showError(type.toUpperCase() + ": " + sensitiveInfo[type][0].value);
                  break;
                }
              }



           // } else {
             // addError('The content of your email is not accessible for policies control.')

            //}


            await context.sync();
            console.log(JSON.stringify(sensitiveInfo));

          });


      };

