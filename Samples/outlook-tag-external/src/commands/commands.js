/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = function (reason) {};

/**
 * Handles the OnMessageRecipientsChanged event.
 * @param event {*} The Office event object
 */
function tagExternal_onMessageRecipientsChangedHandler(event) {
  console.log("tagExternal_onMessageRecipientsChangedHandler method"); //debugging
  console.log("event: " + JSON.stringify(event)); //debugging
  if (event.changedRecipientFields.to) {
    checkForExternalTo();
  }
  if (event.changedRecipientFields.cc) {
    checkForExternalCc();
  }
  if (event.changedRecipientFields.bcc) {
    checkForExternalBcc();
  }
}

/**
 * Determines if there are any external recipients in the To field.
 * If there are, updates the subject of the Outlook message
 * and appends a disclaimer to the message body.
 */
function checkForExternalTo() {
  console.log("checkForExternalTo method"); //debugging

  // Get To recipients.
  console.log("Get To recipients"); //debugging
  Office.context.mailbox.item.to.getAsync(
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get To recipients. " + JSON.stringify(asyncResult.error));
        return;
      }

      const toRecipients = JSON.stringify(asyncResult.value);
      console.log("To recipients: " + toRecipients); //debugging
      if (toRecipients != null
          && toRecipients.length > 0
          && JSON.stringify(toRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        console.log("To includes external users"); //debugging

        // Update sessionData.tagExternalTo.
        Office.context.mailbox.item.sessionData.setAsync(
          "tagExternalTo",
          "true",
          function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("sessionData.setAsync(tagExternalTo) to true succeeded");
          } else {
            console.error("Failed to set tagExternalTo sessionData to true. Error: " + JSON.stringify(asyncResult.error));
            return;
          }
        });
      } else {
        Office.context.mailbox.item.sessionData.setAsync(
          "tagExternalTo",
          "false",
          function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("sessionData.setAsync(tagExternalTo) to false succeeded");
          } else {
            console.error("Failed to set tagExternalTo sessionData to false. Error: " + JSON.stringify(asyncResult.error));
            return;
          }
        });
      }

      _checkForExternal();
    });
}
/**
 * Determines if there are any external recipients in the Cc field.
 * If there are, updates the subject of the Outlook message
 * and appends a disclaimer to the message body.
 */
function checkForExternalCc() {
  console.log("tagExternal method"); //debugging

  // Get Cc recipients.
  console.log("Get Cc recipients"); //debugging
  Office.context.mailbox.item.cc.getAsync(
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get Cc recipients. " + JSON.stringify(asyncResult.error));
        return;
      }
      
      const ccRecipients = JSON.stringify(asyncResult.value);
      console.log("Cc recipients: " + ccRecipients); //debugging
      if (ccRecipients != null
          && ccRecipients.length > 0
          && JSON.stringify(ccRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        console.log("Cc includes external users"); //debugging

        // Update sessionData.tagExternalCc.
        Office.context.mailbox.item.sessionData.setAsync(
          "tagExternalCc",
          "true",
          function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("sessionData.setAsync(tagExternalCc) to true succeeded");
          } else {
            console.error("Failed to set tagExternalCc sessionData to true. Error: " + JSON.stringify(asyncResult.error));
            return;
          }
        });
      } else {
        Office.context.mailbox.item.sessionData.setAsync(
          "tagExternalCc",
          "false",
          function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("sessionData.setAsync(tagExternalCc) to false succeeded");
          } else {
            console.error("Failed to set tagExternalCc sessionData to false. Error: " + JSON.stringify(asyncResult.error));
            return;
          }
        });
      }

      _checkForExternal();
    });
}
/**
 * Determines if there are any external recipients in the Bcc field.
 * If there are, updates the subject of the Outlook message
 * and appends a disclaimer to the message body.
 */
function checkForExternalBcc() {
  console.log("checkForExternalBcc method"); //debugging

  // Get Bcc recipients.
  console.log("Get Bcc recipients"); //debugging
  Office.context.mailbox.item.bcc.getAsync(
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get Bcc recipients. " + JSON.stringify(asyncResult.error));
        return;
      }

      const bccRecipients = JSON.stringify(asyncResult.value);
      console.log("Bcc recipients: " + bccRecipients); //debugging
      if (bccRecipients != null
          && bccRecipients.length > 0
          && JSON.stringify(bccRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        console.log("Bcc includes external users"); //debugging

        // Update sessionData.tagExternalBcc.
        Office.context.mailbox.item.sessionData.setAsync(
          "tagExternalBcc",
          "true",
          function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("sessionData.setAsync(tagExternalBcc) to true succeeded");
          } else {
            console.error("Failed to set tagExternalBcc sessionData to true. Error: " + JSON.stringify(asyncResult.error));
            return;
          }
        });
      } else {
        Office.context.mailbox.item.sessionData.setAsync(
          "tagExternalBcc",
          "false",
          function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("sessionData.setAsync(tagExternalBcc to false) succeeded");
          } else {
            console.error("Failed to set tagExternalBcc sessionData to false. Error: " + JSON.stringify(asyncResult.error));
            return;
          }
        });
      }

      _checkForExternal();
    });
}
/**
 * Checks the session data to determine if any field contains external recipients.
 */
function _checkForExternal() {
  console.log("_checkForExternal method"); //debugging

  // Get sessionData to determine if any fields have external recipients.
  Office.context.mailbox.item.sessionData.getAllAsync(
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get all sessionData. " + JSON.stringify(asyncResult.error));
        return;
      }

      const sessionData = JSON.stringify(asyncResult.value);
      console.log("Current SessionData: " + sessionData); //debugging
      if (sessionData != null
        && sessionData.length > 0
        && sessionData.includes("true")) {
        console.log("At least one recipients field includes external users"); //debugging
        _tagExternal(true);
      } else {
        console.log("There are no external recipients"); //debugging
        _tagExternal(false);
      }
    });
}
/**
 * If there are any external recipients, tags the subject of the Outlook item
 * as "External" and appends a disclaimer to the item body. If there are
 * no external recipients, ensures the tag is not present and clears the disclaimer.
 * @param hasExternal {bool} If there are any external recipients
 */
function _tagExternal(hasExternal) {
  console.log("_tagExternal method"); //debugging

  // External subject tag.
  const externalTag = "[External]";

  if (hasExternal) {
    console.log("External: Get Subject"); //debugging
    // Ensure "[External]" is prepended to the subject.
    Office.context.mailbox.item.subject.getAsync(
      function (asyncResult) {
        // Handle success or error.
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));
          return;
        }

        console.log("Current Subject: " + JSON.stringify(asyncResult.value)); //debugging
        var subject = asyncResult.value;
        if (!subject.includes(externalTag)) {
          subject = `${externalTag} ${subject}`;
          console.log("Updated Subject: " + subject); //debugging
          Office.context.mailbox.item.subject.setAsync(
            subject,
            function (asyncResult) {
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to set Subject. " + JSON.stringify(asyncResult.error));
                return;
              }
          });
        }
    });

    // Set disclaimer as there are external recipients.
    const disclaimer = '<p style = "color:blue"><i>Caution: This email includes external recipients.</i></p>';
    console.log("Set disclaimer"); //debugging
    Office.context.mailbox.item.body.appendOnSendAsync(
      disclaimer,
      {
        "coercionType": Office.CoercionType.Html
      },
      function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to set disclaimer via appendOnSend. " + JSON.stringify(asyncResult.error));
          return;
        }
        console.log("Set disclaimer succeeded"); //debugging
      }
    );
  } else {
    console.log("Internal: Get subject"); //debugging
    // Ensure "[External]" is not part of the subject.
    Office.context.mailbox.item.subject.getAsync(
      function (asyncResult) {
        // Handle success or error.
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));
          return;
        }

        console.log("Current subject: " + JSON.stringify(asyncResult.value)); //debugging
        const currentSubject = asyncResult.value;
        if (currentSubject.startsWith(externalTag)) {
          const updatedSubject = currentSubject.replace(externalTag, "");
          console.log("Updated subject: " + updatedSubject); //debugging
          const subject = updatedSubject.trim();
          console.log("Trimmed subject: " + subject); //debugging
          Office.context.mailbox.item.subject.setAsync(
            subject,
            function (asyncResult) {
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to set subject. " + JSON.stringify(asyncResult.error));
                return;
              }
            });
        }
    });

    // Clear disclaimer as there aren't any external recipients.
    console.log("Clear disclaimer"); //debugging
    Office.context.mailbox.item.body.appendOnSendAsync(
      null,
      function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to clear disclaimer via appendOnSend. " + JSON.stringify(asyncResult.error));
          return;
        }
        console.log("Clear disclaimer succeeded"); //debugging
      }
    );
  }
}

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate("tagExternal_onMessageRecipientsChangedHandler", tagExternal_onMessageRecipientsChangedHandler);
