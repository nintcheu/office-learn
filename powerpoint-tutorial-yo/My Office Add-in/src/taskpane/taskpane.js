/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { base64Image } from "../../base64Image";
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert-text").onclick = () => clearMessage(insertText);
    document.getElementById("validate-text").onclick = () => clearMessage(detectSensitiveData);

    document.getElementById("insert-image").onclick = () => clearMessage(insertImage);

    document.getElementById("get-slide-metadata").onclick = () => clearMessage(getSlideMetadata);
    document.getElementById("add-slides").onclick = () => tryCatch(addSlides);
    document.getElementById("go-to-first-slide").onclick = () => clearMessage(goToFirstSlide);
    document.getElementById("go-to-next-slide").onclick = () => clearMessage(goToNextSlide);
    document.getElementById("go-to-previous-slide").onclick = () => clearMessage(goToPreviousSlide);
    document.getElementById("go-to-last-slide").onclick = () => clearMessage(goToLastSlide);
  }
});

function insertImage() {
  // Call Office.js to insert the image into the document.
  Office.context.document.setSelectedDataAsync(
    base64Image,
    {
      coercionType: Office.CoercionType.Image
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
      }
    }
  );
}

function insertText() {
  Office.context.document.setSelectedDataAsync("John's SSN is 123-45-6789 and his credit card number is 4111111111111111. He was born on 05/12/1980", (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function getSlideMetadata() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    } else {
      setMessage("Metadata for selected slides: " + JSON.stringify(asyncResult.value));
    }
  });
}
  
async function addSlides() {
  await PowerPoint.run(async function (context) {
    context.presentation.slides.add();
    context.presentation.slides.add();

    await context.sync();

    goToLastSlide();
    setMessage("Success: Slides added.");
  });
}

function goToFirstSlide() {
  Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToLastSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToPreviousSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToNextSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

function setMessage(message) {
  document.getElementById("message").innerText = message;
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    document.getElementById("message").innerText = "";
    await callback();
  } catch (error) {
    setMessage("Error: " + error.toString());
  }
}


 async function detectSensitiveData() {
  
  // This example gets the body of the item as plain text.
  await PowerPoint.run(async  (context) => {


    context.presentation.load("slides");
    await context.sync();
    let slides = context.presentation.slides.items;
    console.log("number slides ", slides.length);

    console.log("slides ", slides);
    console.log("slides count ", context.presentation.slides.getCount());

    return;

    //let body = context.document.body;

      //if (result.status == "succeeded") {

        const patterns = {
          ssn: /\b(?!000|666|9\d{2})([0-8]\d{2}|7([0-6]\d|7[012]))([-]?)\d{2}\3\d{4}\b/,
          creditCard: /\b(?:4[0-9]{12}(?:[0-9]{3})?|5[1-5][0-9]{14}|3[47][0-9]{13}|3(?:0[0-5]|[68][0-9])[0-9]{11}|6(?:011|5[0-9]{2})[0-9]{12}|(?:2131|1800|35\d{3})\d{11})\b|\b(?:(?:4[0-9]{3}|5[1-5][0-9]{2}|6[0-9]{3}|3[47][0-9]{2})[- ]?[0-9]{4}[- ]?[0-9]{4}[- ]?[0-9]{4})\b/,
          dateOfBirth: /\b(0[1-9]|1[0-2])[/-](0[1-9]|[12]\d|3[01])[/-](19|20)\d{2}\b/,
          email: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/,
          phoneNumber: /\b(\+\d{1,2}\s?)?1?\-?\.?\s?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}\b/,
          ipAddress: /\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b/,
          passportNumber: /\b[A-Z]{1,2}[0-9]{6,9}\b/,
          driverLicense: /\b[A-Z]{1}[0-9]{5,7}\b/,
          bankAccount: /\b[0-9]{8,17}\b/
        };

        const sensitiveInfo = {};
        let text = result.value;

        for (const [type, pattern] of Object.entries(patterns)) {
          const matches = text.match(pattern);
          if (matches) {
            sensitiveInfo[type] = matches.map(match => ({
              value: match,
              index: text.indexOf(match)
            }));
            addInformational(type.toUpperCase() + " detected as sensitive data");
            showError(type.toUpperCase() + ": " + sensitiveInfo[type][0].value);
            break;
          }
        }



      //} else {
        //addError('The content of your email is not accessible for policies control.')

      //}

    });


}
