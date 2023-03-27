let config = [];

let langBundle = {
    "NO_LOC_FOUND": {
        "en" :"No Location found.",
        "de" :"Kein Standord gefunden."
    },
    "CHANGE_FILTER_CRITERIA": {
        "en" :"Change the location filter parameters.",
        "de" :"Passen Sie Ihre Filterkriterien an."
    },
    "LOC_ADDED": {
        "en" :"Location added.",
        "de" :"Standord hinzugefügt."
    },
}

/**
 * Triggered when Tabpane is ready and loaded in outlook.
 */
Office.onReady((info) => {
  initialize();
});

/**
 * Used to test in browser.
 */
//$(document ).ready(function() {
//  initialize();
//});

/**
 * Initialize list from location configuration json.
 */
function initialize(){
    $.getJSON( "config.json", function( data ) {
        config = data;
        printLocations("");
    });
}

/**
 * Print cards with locations filtered by wildcard filter strings.
 * @param {*} filter 
 */
function printLocations(filter){
    var locList = config.locations.filter(loc => loc.name.toLowerCase().includes(filter.toLowerCase()));
    $("#locations").empty();

    locList.forEach(loc => {
        var card = '<div class="row" style="margin:1em">';
        card += '<div class="card" style="cursor: pointer;" onclick="addLocationRecipient(\'' + loc.email + '\')">';
        card += '<div class="card-content"><span class="card-title">' + loc.name + '</span><p>' + loc.description + '</p></div>';
        card += '</div></div>';
        $("#locations").append(card)
    });

    // no location found
    if(locList.length == 0){
        var card = '<div class="row" style="margin:1em">';
        card += '<div class="card">';
        card += '<div class="card-content"><span class="card-title">' + getLocalizedString("NO_LOC_FOUND") + '</span><p>' + getLocalizedString("CHANGE_FILTER_CRITERIA") + '</p></div>';
        card += '</div></div>';
        $("#locations").append(card)
    }
}

/**
 * Adds the location email as recipient to the appointment.
 * @param {*} locEmail 
 */
function addLocationRecipient(locEmail) {
    showMessage(locEmail)
    Office.context.mailbox.item.requiredAttendees.addAsync(
        [{
            "emailAddress": locEmail
        }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else{
               showMessage(getLocalizedString("LOC_ADDED")); 
            }
        });
}

function showMessage(msg) {
     $("#banner").append('<div class="teal lighten-2" style="padding: 0.5em; margin:1em">' + msg + '</div>');
    setTimeout(function() { $("#banner").empty()}, 3000);
}

function getLocalizedString(key) {
    var displayLanguage;
    try {
        displayLanguage = Office.context.displayLanguage;
    } catch(e) {
        displayLanguage = "en-us";
    }
    var lng = displayLanguage.toLowerCase().startsWith("de-") ? "de" : "en";

    return langBundle[key][lng];
}