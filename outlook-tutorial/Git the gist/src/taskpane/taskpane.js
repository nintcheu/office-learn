(function () {
  'use strict';

  let config;
  let settingsDialog;

  Office.initialize = function (reason) {

    jQuery(document).ready(function () {

      config = getConfig();

      // Check if add-in is configured.
      if (config && config.gitHubUserName) {
        // If configured, load the gist list.
        loadGists(config.gitHubUserName);
      } else {
        // Not configured yet.
        $('#not-configured').show();
      }

      // When insert button is selected, build the content
      // and insert into the body.
      $('#insert-button').on('click', function () {
        const gistId = $('.ms-ListItem.is-selected').val();
        getGist(gistId, function (gist, error) {
          if (gist) {
            buildBodyContent(gist, function (content, error) {
              if (content) {
                Office.context.mailbox.item.body.setSelectedDataAsync(content,
                  { coercionType: Office.CoercionType.Html }, function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      showError('Could not insert gist: ' + result.error.message);
                    }
                  });
              } else {
                showError('Could not create insertable content: ' + error);
              }
            });
          } else {
            showError('Could not retrieve gist: ' + error);
          }
        });
      });


      //When verify-content-button is selected, get the content of the email body
      $('#verify-content-button').on('click', function () {
        //addInformational("Any sensitive data found in your mail. Great job! You respect our policies.");
        //showError("You found sensitive date in your mail. You don't respect the policies. Please, take action to correct the issue.");

        // This example gets the body of the item as plain text.
        Office.context.mailbox.item.body.getAsync(
          "text",
          { asyncContext: "This is passed to the callback" },
          function callback(result) {

            if (result.status == "succeeded") {

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



            } else {
              addError('The content of your email is not accessible for policies control.')

            }

          });




      });


      function addInformational(msgInfo) {
        // Adds a persistent information notification to the mail item.
        const id = $("#notificationId").val().toString();
        const details =
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: msgInfo,
          icon: "icon1",
          persistent: true
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync(id, details, function (result) {

        });
      }

      function getEmailBody() {


        if (sizeof(sensitiveInfo.ssn)) {
          showError("SSN detected: " + sensitiveInfo.ssn[0].value);

        }
        if (sizeof(sensitiveInfo.creditCard)) {
          showError("SSN detected: " + sensitiveInfo.creditCard[0].value);


        }
        if (sizeof(sensitiveInfo.dateOfBirth)) {
          showError("SSN detected: " + sensitiveInfo.dateOfBirth[0].value);
        }



      }

      function addError(errMsg) {
        // Adds an error notification to the mail item.
        const id = $("#notificationId").val().toString();
        const details =
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: errMsg
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync(id, details, function (result) {

        });
      }

      /**
       * 
       * @param {string} text 
       * 
       * @returns JSON
       * {
            ssn: [{ value: '123-45-6789', index: 13 }],
            creditCard: [{ value: '4111111111111111', index: 47 }],
            dateOfBirth: [{ value: '05/12/1980', index: 86 }]
          } 
       */
      function detectSensitiveInfo(text) {
        const patterns = {
          ssn: /\b(?!000|666|9\d{2})([0-8]\d{2}|7([0-6]\d|7[012]))([-]?)\d{2}\3\d{4}\b/,
          creditCard: /\b(?:4[0-9]{12}(?:[0-9]{3})?|5[1-5][0-9]{14}|3[47][0-9]{13}|3(?:0[0-5]|[68][0-9])[0-9]{11}|6(?:011|5[0-9]{2})[0-9]{12}|(?:2131|1800|35\d{3})\d{11})\b/,
          dateOfBirth: /\b(0[1-9]|1[0-2])[/-](0[1-9]|[12]\d|3[01])[/-](19|20)\d{2}\b/
        };

        const results = {};

        for (const [type, pattern] of Object.entries(patterns)) {
          const matches = text.match(pattern);
          if (matches) {
            results[type] = matches.map(match => ({
              value: match,
              index: text.indexOf(match)
            }));
          }
        }

        return results;
      }

      // When the settings icon is selected, open the settings dialog.
      $('#settings-icon').on('click', function () {
        // Display settings dialog.
        let url = new URI('dialog.html').absoluteTo(window.location).toString();
        if (config) {
          // If the add-in has already been configured, pass the existing values
          // to the dialog.
          url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
        }

        const dialogOptions = { width: 20, height: 40, displayInIframe: true };

        Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
          settingsDialog = result.value;
          settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
          settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        });
      })
    });
  };

  function loadGists(user) {
    $('#error-display').hide();
    $('#not-configured').hide();
    $('#gist-list-container').show();

    getUserGists(user, function (gists, error) {
      if (error) {

      } else {
        $('#gist-list').empty();
        buildGistList($('#gist-list'), gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    $('#insert-button').removeAttr('disabled');
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
  }

  function showError(error) {
    $('#not-configured').hide();
    $('#gist-list-container').hide();
    $('#error-display').text(error);
    $('#error-display').show();
  }


  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function (result) {
      settingsDialog.close();
      settingsDialog = null;
      loadGists(config.gitHubUserName);
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();