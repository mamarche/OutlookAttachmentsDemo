/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
        });

        $("#testGetAttachments").click(GetAttachmentsInfo);
    };

    function GetAttachmentsInfo() {
        //richiamo l'API per ottenere le informazioni sugli allegati della mail
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }

    function attachmentTokenCallback(asyncResult, userContext) {
        //creo un oggetto per memorizzare i dati degli allegati
        //da passare al servizio
        var serviceRequest = new Object();
        serviceRequest.attachmentToken = "";
        serviceRequest.ewsUrl = "";
        serviceRequest.state = -1;
        serviceRequest.attachments = new Array();

        if (asyncResult.status === "succeeded") {
            //valorizzo le proprietà (token, url e info sugli allegati)
            serviceRequest.attachmentToken = asyncResult.value;
            serviceRequest.state = 3;
            serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
            serviceRequest.attachments = new Array();
            for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
                serviceRequest.attachments[i] = JSON.parse(JSON.stringify(Office.context.mailbox.item.attachments[i]._data$p$0));
            }

            //richiamo il metodo del controller passando l'oggetto serviceRequest
            $.ajax({
                url: 'https://outlookattachmentsdemoservice.azurewebsites.net/api/attachments',
                type: 'POST',
                contentType: 'application/json; charset=utf-8',
                cache: false,
                data: JSON.stringify(serviceRequest)
            }).done(function (response) {
                var names = "";
                for (var v = 0; v < response.attachmentNames.length; v++) {
                    names = names + ", " + response.attachmentNames[v];
                }

                showNotification("Attachments processed", "Number of attachments: " + response.attachmentsProcessed +
                                                           " (" + names + ")");
            }).fail(function (status) {
                showNotification("Error", status.statusText);
            });

        } else {
            showNotification("Error", "Could not get callback token: " + asyncResult.error.message);
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();