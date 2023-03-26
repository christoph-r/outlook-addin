    Office.onReady((info) => { });

    function addLocation(locEmail) {
        Office.context.mailbox.item.requiredAttendees.addAsync(
            [{
                "displayName": locEmail, 
                "emailAddress": locEmail
            }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    write(asyncResult.error.message);
                }
            });
    }
