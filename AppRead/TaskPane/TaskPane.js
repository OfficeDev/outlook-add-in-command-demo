// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            loadProps(); //calls each of the methods to load properties within each section
        });
    };

    function loadBodyProps(item) {

        $('#message-props').show();
        loadReportId(item);
        loadSender(item);
        loadRecipientsTo(item);
        loadRecipientsCC(item);
        loadRecipientsBCC(item);
        loadMessageSubject(item);
        loadDateSent(item);
        loadMessageId(item);
    }

    function loadDetectionDetails(item) {
        loadInitialSev(item);
        loadOverride(item);
        loadOverrideJustification(item);
        loadFalsePositive(item);
    }

    function loadRuleHitDetails(item) {
        $('#rule-hit-props').show();
        getMatchesInGIR(item);
    }

    //*************************Message Details Methods*********************


    function loadReportId(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                // i => ignore case
                // g => global match, i.e., doesn't stop after first match
                var regex = new RegExp('(?:Report Id: )(\\S+)', 'ig');
                var match; //= new Array();
                match = regex.exec(bodyText);
                if (match != null) {
                    //var divElement = document.getElementById('reportId');
                    //divElement.innerHTML = divElement.innerHTML + " \"" + match[1] + "\""; //we use match[1] because match[0] returns the entire matched string including non-capturing groups
                    $('#reportId').text(match[1]);
                }
                else {
                    $('#reportId').text('--');
                }

            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }


    /*
    Get the sender of the message detailed in the incident report (not the gir itself)
    */
    function loadSender(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;

                var regex = new RegExp('(?:Sender: )(\\S+) (\\S+) (\\S+)', 'ig');
                var match;
                match = regex.exec(bodyText);
                if (match != null) {
                    $('#messageSender').text(match[1] + ' ' + match[2] + ' ' + match[3]);

                }
                else {
                    $('#messageSender').text('--');
                }

            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }

    /*
   Get the 'To Recipients' of the message detailed in the incident report (not the gir itself)
   */
    function loadRecipientsTo(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:To: )(\\S+) (\\S+) (\\S+)|(?:To: )(\\S+)', 'g');
                var matches = getAllMatches(regex, bodyText); //array of matches

                if (matches != null) {
                    $('#recipientsTo').text(buildEmailAddressString(matches));
                }
                else {
                    $('#recipientsTo').text('--');
                }

            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }


    /*
    Get the 'CC Recipients' of the message detailed in the incident report (not the gir itself)
    */
    function loadRecipientsCC(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:CC: )(\\S+) (\\S+) (\\S+)|(?:CC: )(\\S+)', 'g');
                var matches = getAllMatches(regex, bodyText); //array of matches

                if (matches != null) {
                    $('#recipientsCC').text(buildEmailAddressString(matches));
                }
                else {
                    $('#recipientsCC').text('--');
                }

            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }

    /*
    Get the 'CC Recipients' of the message detailed in the incident report (not the gir itself)
    */
    function loadRecipientsBCC(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:BCC: )(\\S+) (\\S+) (\\S+)|(?:BCC: )(\\S+)', 'g');

                var matches = getAllMatches(regex, bodyText); //array of matches

                if (matches != null) {
                    $('#recipientsBCC').text(buildEmailAddressString(matches));
                }
                else {
                    $('#recipientsBCC').text('--');
                }

            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }

    //bug, need to figure out how to just get the subject on it's own reliablely
    function loadMessageSubject(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {

                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:Subject: )(.*?)(?:Recipients: )', 'gi');   //super duper hacky but that's alright for now until we use html
                var match; //= new Array();
                match = regex.exec(bodyText);
                if (match != null) {
                    $('#messageSubject').text(match[1]); 

                }
                else {
                    $('#messageSubject').text('--');
                }
            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }

    function loadMessageId(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:Message Id: <)(\\S+)(?:>)', 'ig');
                var match; //= new Array();
                match = regex.exec(bodyText);
                if (match != null) {
                    $('#messageId').text(match[1]);

                }
                else {
                    $('#messageId').text('--');
                }
            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }

    function loadDateSent(item) {
                    $('#dateSent').text('Date/Time is not currently available in the text of this document');
    }


    //*********************************Detection Details Methods*****************

    function loadInitialSev(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:Severity: )(\\S+)', 'ig');
                var match; //= new Array();
                match = regex.exec(bodyText);
                if (match != null) {
                    $('#initialSeverity').text(match[1]);

                }
                else {
                    $('#initialSeverity').text('--');
                }
            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }

    function loadOverride(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:Override: )(\\S+)', 'ig');
                var match; //= new Array();
                match = regex.exec(bodyText);
                if (match != null) {
                    $('#override').text(match[1]);

                }
                else {
                    $('#override').text('--');
                }
            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }

    function loadOverrideJustification(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:Justification: )(\\S+)', 'ig');  //this is just a guess, it could be labeled as something completely different, I just haven't seen an example yet
                var match; //= new Array();
                match = regex.exec(bodyText);
                if (match != null) {
                    $('#overrideJustification').text(match[1]);

                }
                else {
                    $('#overrideJustification').text('No override action was taken');
                }
            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }

    function loadFalsePositive(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;
                var regex = new RegExp('(?:False Positive: )(\\S+)', 'ig');
                var match; //= new Array();
                match = regex.exec(bodyText);
                if (match != null) {
                    $('#falsePositive').text(match[1]);

                }
                else {
                    $('#falsePositive').text('--');
                }
            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }


    //*******************************Generate rule matches****************

    function getMatchesInGIR(item) {
        if (Office.context.mailbox.item.body.getAsync !== undefined) {
            Office.context.mailbox.item.body.getAsync('text', function (asyncResult) {
                var bodyText = asyncResult.value;


                var dataClassificationRegex = new RegExp('(?:Data Classification: )(.*?)(?:,)', 'ig');
                var countRegex = new RegExp('(?:Count: )(\\S+)(?:,)', 'ig');
                var ruleRegex = new RegExp('(?:Rule Hit: )(.*?)(?:,)', 'ig');
                var actionRegex = new RegExp('(?:False Positive: )(\\S+)', 'ig');

                var dataClassificationResult = getAllMatches(dataClassificationRegex, bodyText);
                var countResult = getAllMatches(countRegex, bodyText);
                var ruleResult = getAllMatches(ruleRegex, bodyText);
                var actionResult = getAllMatches(actionRegex, bodyText);
             

                //making the (perhaps incorrect) assumption that there will be an equal number of rules, counts, actions, and classifications
                //additionally, all matches above only have one capture group, hence the count+=2
                var count = 1;
                if (ruleResult[count-1] != null) {
                        $('#rule1').text(ruleResult[count]);
                        $('#dataClassification1').text(dataClassificationResult[count]);
                        $('#count1').text(countResult[count]);
                        $('#action1').text(actionResult[count]);
                }
                else {
                    //shouldn't hit here because a rule will always trigger the gir but just to be safe...
                    $('#rule1').text('--');
                    $('#dataClassification1').text('--');
                    $('#count1').text('--');
                    $('#action1').text('--');
                }

            });
        }
        else { // Method not available
            app.showNotification('Warning', 'The body.getAsync() method is not available in this version of Outlook. Body parsing was skipped');
        }
    }



    //********************************Helper methods********************

    //return the list of addresses
    function buildEmailAddressString(arrayOfMatches) {
        var resultString = '';
        //app.showNotification('Warning', arrayOfMatches[0] + ' 1 ' + arrayOfMatches[1] + ' 2 ' + arrayOfMatches[2] + ' 3 ' +arrayOfMatches[3] + ' 4 ' +arrayOfMatches[4] + ' ');
        for (var count = 0; count < arrayOfMatches.length; count += 5) {
            if (arrayOfMatches[count + 1] != null) {
                resultString = resultString + arrayOfMatches[count + 1] + ' ' + arrayOfMatches[count + 2] + ' ' + arrayOfMatches[count + 3] + '; ';
            }
            else {
                resultString = resultString + arrayOfMatches[count + 4] + '; ';
            }
        }

        return resultString;
    }
    
    /*
    *  input a regex that searches for *all* of the occurances (so add g)
    *  bodyText should be in text form
    *  returns an array of the matches
    */
    function getAllMatches(regex, bodyText) {

        var match;
        var result = [];

        while ((match = regex.exec(bodyText)) !== null) {
            result = result.concat(match);
        }

        if (result.length != 0) {
            return result;
        }
        else {
            return null;
        }
    }

    /*
    * input: result of all regex, an array with the position within a single match
    * which should be returned
    * returns: array with only desired strings
    */
    function getTextFromRegexMatch(resultArray, arrayOfMatchPositionsToReturn, lengthOfMatchPositions) {



    }



    //*************************************Final Load Method***************

    // Load properties from the Item base object, then load the
    // type-specific properties.
    function loadProps() {
        var item = Office.context.mailbox.item;
        if (item.itemType == Office.MailboxEnums.ItemType.Message) {
            loadBodyProps(item);
            loadDetectionDetails(item);
            loadRuleHitDetails(item);
        }
        else {
            loadAppointmentProps(item);
            app.showNotification('Warning', 'This iteration is not supposed to run for appointments.');
        }
        $('#body')
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