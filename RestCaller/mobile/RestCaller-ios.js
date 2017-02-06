$(document).ready(function() {
  
});

// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

(function () {
  'use strict';

  var myApp;

  var rawToken = '';
  var parsedToken = '';

  var getItemSpinnerElement = null;
  var getItemSpinner = null;
  var updateItemSpinnerElement = null;
  var updateItemSpinner = null;

  var markUnreadPayload = { IsRead: false };
  var flagFollowupPayload = { Flag: { FlagStatus: 'Flagged' } };
  var applyCategoryPayload = { Categories: ['Red Category'] };

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
      $(document).ready(function () {
          myApp = new Framework7();

          var getItemView = myApp.addView('#get-item-view');
          var getItemView = myApp.addView('#update-item-view');
          var getItemView = myApp.addView('#item-details-view');

          var changePicker = myApp.picker({
            input: '#change-picker',
            cols: [
              {
                textAlign: 'center',
                values: [
                  'Mark unread',
                  'Flag for followup',
                  'Apply category'
                ]
              }
            ]
          });

          $('#change-picker').change(function() {
            var newValue = $('#change-picker').val();
            loadItemChangePayload(newValue);
          });

          $('#parse-token-toggle').change(function() {
            loadToken($('#parse-token-toggle').is(':checked'));
          });

          $('.get-item-button').click(function() {
            getItemViaRest();
          });

          $('.update-item-button').click(function() {
            updateItemViaRest();
          });

          loadRestDetails();
      });
  };

  function loadRestDetails() {
    $('.hostname').text(Office.context.mailbox.diagnostics.hostName);
    $('.hostversion').text(Office.context.mailbox.diagnostics.hostVersion);
    $('.owaview').text(Office.context.mailbox.diagnostics.OWAView);

    myApp.showPreloader();

    var restId = '';
    if (Office.context.mailbox.diagnostics.hostName !== 'OutlookIOS') {
      // Loaded in non-mobile context, so ID needs to be converted
      restId = Office.context.mailbox.convertToRestId(
        Office.context.mailbox.item.itemId,
        Office.MailboxEnums.RestVersion.Beta
      );
    } else {
      restId = Office.context.mailbox.item.itemId;
    }

    // Build the URL to the item
    var itemUrl = Office.context.mailbox.restUrl + 
      '/api/beta/me/messages/' + restId;

    $('.resturl-display code').text(itemUrl);
    
    Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
      myApp.hidePreloader();
      if (result.status === "succeeded") {
        rawToken = result.value;
        loadToken($('#parse-token-toggle').is(':checked'));
        enableButtons();
      } else {
        rawToken = 'error';
      }
    });
  }

  function loadToken(parseToken) {
    var code = $('.token-display code');
    if (rawToken === 'error') {
      code.text('ERROR RETRIEVING TOKEN');
      return;
    }

    if (parseToken) {
      if (parsedToken === '') {
        parsedToken = jwt_decode(rawToken);
      }

      code.text(JSON.stringify(parsedToken, null, 2));
    } else {
      code.text(rawToken);
    }
  }

  function getItemViaRest() {
    var itemUrl = $('.resturl-display code').text();

    myApp.showPreloader();
    
    $.ajax({
      url: itemUrl,
      dataType: 'json',
      headers: { 'Authorization': 'Bearer ' + rawToken }
    }).done(function(item){
      myApp.hidePreloader();
      $('.item-display code').text(
        JSON.stringify(item, null, 2)
      );
    }).fail(function(error){
      myApp.hidePreloader();
      $('.item-display code').text(JSON.stringify(error, null, 2));
    });
  }

  function updateItemViaRest() {
    var itemUrl = $('.resturl-display code').text();
    var payload = $('.update-display code').text();

    myApp.showPreloader();
    
    $.ajax({
      type: 'PATCH',
      url: itemUrl,
      dataType: 'json',
      data: payload,
      headers: { 
        'Authorization': 'Bearer ' + rawToken,
        'Content-Type': 'application/json' 
      }
    }).done(function(item){
      myApp.hidePreloader();
      $('.update-display code').text(
        JSON.stringify(item, null, 2)
      );
    }).fail(function(error){
      myApp.hidePreloader();
      $('.update-display code').text(JSON.stringify(error, null, 2));
    });
  }

  function loadItemChangePayload(payloadName) {
    $('.update-display code').text('loadpayload');
    var payloadText = '';

    switch(payloadName) {
      case "Mark unread":
        payloadText = JSON.stringify(markUnreadPayload, null, 2);
        break;
      case "Flag for followup":
        payloadText = JSON.stringify(flagFollowupPayload, null, 2);
        break;
      case "Apply category":
        payloadText = JSON.stringify(applyCategoryPayload, null, 2);
        break;
      default:
        payloadText = "Choose a change..."
    }

    $('.update-display code').text(payloadText);
  }

  function enableButtons() {
    $('.get-item-button').removeClass('disabled');
    $('.update-item-button').removeClass('disabled');
  }
})();

// MIT License: 
 
// Permission is hereby granted, free of charge, to any person obtaining 
// a copy of this software and associated documentation files (the 
// ""Software""), to deal in the Software without restriction, including 
// without limitation the rights to use, copy, modify, merge, publish, 
// distribute, sublicense, and/or sell copies of the Software, and to 
// permit persons to whom the Software is furnished to do so, subject to 
// the following conditions: 
 
// The above copyright notice and this permission notice shall be 
// included in all copies or substantial portions of the Software. 
 
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.