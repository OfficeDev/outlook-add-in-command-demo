// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="../App.js" />

(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
      $(document).ready(function () {
          app.initialize();

          if (isPersistenceSupported()) {
            // Set up ItemChanged event
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem);
          }

          loadProps(Office.context.mailbox.item);
      });
  };

  function isPersistenceSupported() {
    // This feature is part of the preview 1.5 req set
    // Since 1.5 isn't fully implemented, just check that the 
    // method is defined.
    // Once 1.5 is implemented, we can replace this with
    // Office.context.requirements.isSetSupported('Mailbox', 1.5)
    return Office.context.mailbox.addHandlerAsync !== undefined;
  };

  function loadNewItem(eventArgs) {
    loadProps(Office.context.mailbox.item);
  };
  
  // Take an array of AttachmentDetails objects and
  // build a list of attachment names, separated by a line-break
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }
      
      return returnString;
    }
    
    return "None";
  }
  
  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }
  
  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }
      
      return returnString;
    }
    
    return "None";
  }
  
  // Load properties from a Message object
  function loadMessageProps(item) {
    $('#message-props').show();
    
    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
  }
  
  // Load properties from an Appointment object
  function loadAppointmentProps(item) {
    $('#appointment-props').show();
    
    $('#appt-attachments').html(buildAttachmentsString(item.attachments));
    $('#end').text(item.end.toLocaleString());
    $('#location').text(item.location);
    $('#appt-normalizedSubject').text(item.normalizedSubject);
    $('#optionalAttendees').html(buildEmailAddressesString(item.optionalAttendees));
    $('#organizer').html(buildEmailAddressString(item.organizer));
    $('#requiredAttendees').html(buildEmailAddressesString(item.requiredAttendees));
    $('#resources').html(buildEmailAddressesString(item.resources));
    $('#start').text(item.start.toLocaleString());
    $('#appt-subject').text(item.subject);
  }
  
  // Load properties from the Item base object, then load the
  // type-specific properties.
  function loadProps(item) {
    
    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);
    
    item.body.getAsync('html', function(result){
      if (result.status === 'succeeded') {
        $('#bodyHtml').text(result.value);
      }
    });
    
    item.body.getAsync('text', function(result){
      if (result.status === 'succeeded') {
        $('#bodyText').text(result.value);
      }
    });
    
    if (item.itemType == Office.MailboxEnums.ItemType.Message){
      loadMessageProps(item);
    }
    else {
      loadAppointmentProps(item);
    }
  }
})();