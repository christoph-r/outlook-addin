Office.onReady((info) => {
});

// Add location to recipients
function addLocation(event) {
    Office.context.mailbox.item.requiredAttendees.addAsync(
        [{
            "displayName": "Pansy Valenzuela",
            "emailAddress": "pansy@contoso.com"
        }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }
             event.completed();
        }); // End addAsync.
    }
}

// You must register the function with the following line.
Office.actions.associate("addLocation", addLocation);
