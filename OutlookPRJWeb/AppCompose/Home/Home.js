/// <reference path="../App.js" />

(function () {
    'use strict';

    // Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#set-subject').click(setSubject);
            $('#get-subject').click(getSubject);
            $('#add-to-recipients').click(addToRecipients);
        });
    };

    function setSubject() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync("Hello World!");
    }

    function getSubject() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.getAsync(function (result) {
            app.showNotification('Das aktuelle Thema ist', result.value)
        });
    }

    function addToRecipients() {
        var item = Office.context.mailbox.item;
        var addressToAdd = {
            displayName: Office.context.mailbox.userProfile.displayName,
            emailAddress: Office.context.mailbox.userProfile.emailAddress
        };
 
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
        }
    }

})();