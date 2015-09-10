/// <reference path="../App.js" />

(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
      $(document).ready(function () {
          app.initialize();

          loadProps();
      });
  };
  
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
  
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }
  
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
  
  function loadProps() {
    var item = Office.context.mailbox.item;
    
    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);
    
    if (item.itemType == Office.MailboxEnums.ItemType.Message){
      loadMessageProps(item);
    }
    else {
      loadAppointmentProps(item);
    }
  }
})();