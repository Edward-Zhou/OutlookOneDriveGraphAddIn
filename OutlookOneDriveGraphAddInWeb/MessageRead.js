(function () {
  "use strict";

    var messageBanner;
    var overlay;
    var spinner;
    var authenticator;


  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
        // For auth helper
        if (OfficeHelpers.Authenticator.isAuthDialog()) return;

        
        authenticator = new OfficeHelpers.Authenticator();
        authenticator.endpoints.registerMicrosoftAuth(authConfig.clientId, {
            redirectUrl: authConfig.redirectUrl,
            scope: authConfig.scopes
        });

        initializePane();

    });
   };
    
    function initializePane() {
        var username = Office.context.mailbox.userProfile.displayName;
        $("#username").text(username);
        // First attempt to get an SSO token
        if (Office.context.auth !== undefined && Office.context.auth.getAccessTokenAsync !== undefined) {
            Office.context.auth.getAccessTokenAsync(function (result) {
                if (result.status === "succeeded") {
                    // No need to prompt user, use this token to call Web API
                    saveAttachmentsWithSSO(result.value, 1);
                } else if (result.error.code == 13007 || result.error.code == 13005) {
                    // These error codes indicate that we need to prompt for consent
                    Office.context.auth.getAccessTokenAsync({ forceConsent: true }, function (result) {
                        if (result.status === "succeeded") {
                            saveAttachmentsWithSSO(result.value, 1);
                        } else {
                            // Could not get SSO token, proceed with authentication prompt
                            saveAttachmentsWithPrompt(1);
                        }
                    });
                } else {
                    // Could not get SSO token, proceed with authentication prompt
                    saveAttachmentsWithPrompt(attachmentIds);
                }
            });
        }
        else {
            // SSO not supported
            saveAttachmentsWithPrompt(1);
        }
    }
      function saveAttachments(attachmentIds) {
            showSpinner();

            // First attempt to get an SSO token
            if (Office.context.auth !== undefined && Office.context.auth.getAccessTokenAsync !== undefined) {
                Office.context.auth.getAccessTokenAsync(function (result) {
                    if (result.status === "succeeded") {
                        // No need to prompt user, use this token to call Web API
                        saveAttachmentsWithSSO(result.value, attachmentIds);
                    } else if (result.error.code == 13007 || result.error.code == 13005) {
                        // These error codes indicate that we need to prompt for consent
                        Office.context.auth.getAccessTokenAsync({ forceConsent: true }, function (result) {
                            if (result.status === "succeeded") {
                                saveAttachmentsWithSSO(result.value, attachmentIds);
                            } else {
                                // Could not get SSO token, proceed with authentication prompt
                                saveAttachmentsWithPrompt(attachmentIds);
                            }
                        });
                    } else {
                        // Could not get SSO token, proceed with authentication prompt
                        saveAttachmentsWithPrompt(attachmentIds);
                    }
                });
            } else {
                // SSO not supported
                saveAttachmentsWithPrompt(attachmentIds);
            }
        }

      function saveAttachmentsWithSSO(accessToken, attachmentIds) {
            var saveAttachmentsRequest = {
                attachmentIds: attachmentIds,
                messageId: getRestId(Office.context.mailbox.item.itemId)
            };

            $.ajax({
                type: "POST",
                url: "/api/SaveAttachments",
                headers: {
                    "Authorization": "Bearer " + accessToken
                },
                data: JSON.stringify(saveAttachmentsRequest),
                contentType: "application/json; charset=utf-8"
            }).done(function (data) {
                showNotification("Success", "Attachments saved");
            }).fail(function (error) {
                showNotification("Error saving attachments", error.status);
            }).always(function () {
                hideSpinner();
            });
        }

      function saveAttachmentsWithPrompt(attachmentIds) {
            authenticator
                .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft, true)
                .then(function (token) {
                    console.log(token);
                    // Get callback token, which grants read access to the current message
                    // via the Outlook API
                    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                        if (result.status === "succeeded") {
                            var saveAttachmentsRequest = {
                                attachmentIds: attachmentIds,
                                messageId: getRestId(Office.context.mailbox.item.itemId),
                                outlookToken: result.value,
                                outlookRestUrl: getRestUrl(),
                                oneDriveToken: token.access_token
                            };

                            $.ajax({
                                type: "POST",
                                url: "/api/SaveAttachments",
                                data: JSON.stringify(saveAttachmentsRequest),
                                contentType: "application/json; charset=utf-8"
                            }).done(function (data) {
                                showNotification("Success", "Attachments saved");
                            }).fail(function (error) {
                                showNotification("Error saving attachments", error.status);
                            }).always(function () {
                                hideSpinner();
                            });
                        } else {
                            showNotification("Error getting callback token", JSON.stringify(result));
                            hideSpinner();
                        }
                    });
                })
                .catch(function (error) {
                    showNotification("Error authenticating to OneDrive", error);
                    hideSpinner();
                });
        }

      // Helper function for displaying notifications
      function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
      }
})();