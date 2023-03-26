// Initialize the Office Add-in.
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

// The command function.
async function addLocation(event) {

    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
               // Calling event.completed is required. event.completed lets the platform know that processing has completed.
               event.completed();
            }
        }); // End addAsync.
    }
}

// You must register the function with the following line.
Office.actions.associate("addLocation", highlightSelection);
