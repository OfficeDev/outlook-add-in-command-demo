// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="../App.js" />

(function () {
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      app.initialize();

      $('#insertDefault').click(insertDefault);
      $('#insertMsg1').click(insertMsg1);
      $('#insertMsg2').click(insertMsg2);
      $('#insertMsg3').click(insertMsg3);
      $('#insertCustom').click(insertCustom);
    });
  };
  
  function insertText(textToInsert) {
    // Insert as plain text (CoercionType.Text)
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert, 
      { coercionType: Office.CoercionType.Text }, 
      function (asyncResult) {
        // Display the result to the user
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
          app.showNotification("Success", "\"" + textToInsert + "\" inserted successfully.");
        }
        else {
          app.showNotification("Error", "Failed to insert \"" + textToInsert + "\": " + asyncResult.error.message);
        }
      });
  }

  function insertDefault() {
    insertText("Inserted by the Add-in Command Demo add-in.");
  }
  
  function insertMsg1() {
    insertText("Hello World!");
  }
  
  function insertMsg2() {
    insertText("Add-in commands are cool!");
  }
  
  function insertMsg3() {
    insertText("Visit https://developer.microsoft.com/en-us/outlook/ today for all of your add-in development needs.");
  }
  
  function insertCustom() {
    var textToInsert = $('#textToInsert').val();
    insertText(textToInsert);
  }
})();