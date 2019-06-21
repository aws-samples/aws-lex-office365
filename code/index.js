// JavaScript source code

'use strict';

/**
 * This code sample demonstrates an implementation of the Lex Code Hook Interface
 * in order to serve a bot which manages dentist appointments.
 * Bot, Intent, and Slot models which are compatible with this sample can be found in the Lex Console
 * as part of the 'MakeAppointment' template.
 *
 * For instructions on how to set up and test this bot, as well as additional samples,
 *  visit the Lex Getting Started documentation.
 */
var AWS = require("aws-sdk");
var Request = require("request");

// --------------- Helpers to build responses which match the structure of the necessary dialog actions -----------------------

function elicitSlot(sessionAttributes, intentName, slots, slotToElicit, message, responseCard) {
    //console.log("1");
    // sendanote(message);
    return {
        sessionAttributes,
        dialogAction: {
            type: 'ElicitSlot',
            intentName,
            slots,
            slotToElicit,
            message,
            responseCard,
        },
    };
}

function confirmIntent(sessionAttributes, intentName, slots, message, responseCard) {
    //console.log("2");
    //sendanote(message);
    return {
        sessionAttributes,
        dialogAction: {
            type: 'ConfirmIntent',
            intentName,
            slots,
            message,
            responseCard,
        },
    };
}

function close(sessionAttributes, fulfillmentState, message, responseCard) {
    //console.log("3");
    return {
        sessionAttributes,
        dialogAction: {
            type: 'Close',
            fulfillmentState,
            message,
            responseCard,
        },
    };
}

function delegate(sessionAttributes, slots) {
    //console.log("4");

    return {
        sessionAttributes,
        dialogAction: {
            type: 'Delegate',
            slots,
        },
    };
}

// Build a responseCard with a title, subtitle, and an optional set of options which should be displayed as buttons.
function buildResponseCard(title, subTitle, options) {
    //console.log("5");
    let buttons = null;
    if (options != null) {
        buttons = [];
        for (let i = 0; i < Math.min(5, options.length); i++) {
            buttons.push(options[i]);
        }
    }
    return {
        contentType: 'application/vnd.amazonaws.card.generic',
        version: 1,
        genericAttachments: [{
            title,
            subTitle,
            buttons,
        }],
    };
}

// ---------------- Helper Functions --------------------------------------------------

function parseLocalDate(date) {
    /**
     * Construct a date object in the local timezone by parsing the input date string, assuming a YYYY-MM-DD format.
     * Note that the Date(dateString) constructor is explicitly avoided as it may implicitly assume a UTC timezone.
     */
    const dateComponents = date.split(/\-/);
    return new Date(dateComponents[0], dateComponents[1] - 1, dateComponents[2]);
}

function isValidDate(date) {
    try {
        return !(isNaN(parseLocalDate(date).getTime()));
    } catch (err) {
        return false;
    }
}

function incrementTimeByThirtyMins(time) {
    if (time.length !== 5) {
        // Not a valid time
    }
    const hour = parseInt(time.substring(0, 2), 10);
    const minute = parseInt(time.substring(3), 10);
    return (minute === 30) ? `${hour + 1}:00` : `${hour}:30`;
}

// Returns a random integer between min (included) and max (excluded)
function getRandomInt(min, max) {
    const minInt = Math.ceil(min);
    const maxInt = Math.floor(max);
    return Math.floor(Math.random() * (maxInt - minInt)) + minInt;
}


var ssmendpoint = "https://secretsmanager.us-east-1.amazonaws.com";
var region = "us-east-1";
var ssmclient = new AWS.SecretsManager({
    endpoint: ssmendpoint,
    region: region
});

var ssmSecret = "";
var ADDirectoryId = ""; //Consumer Key
var ClientId = ""; //Consumer Secret
var RedirectUri = "";
var ClientSecret = "";
var InvestmentAgentUserId = "";
var PersonalAgentUserId = "";
var o365APISecrets = "";

var getO365APISecretsAndOrchestrate = function (date, callback, intentRequest, bookAppointment) {

    ssmclient.getSecretValue({ "SecretId": o365APISecrets }, function (err, data) {
        if (err) {
            if (err.code === 'ResourceNotFoundException')
                console.log("The requested secret " + o365APISecrets + " was not found");
            else if (err.code === 'InvalidRequestException')
                console.log("The request was invalid due to: " + err.message);
            else if (err.code === 'InvalidParameterException')
                console.log("The request had invalid params: " + err.message);
        }
        else {
            if (data.SecretString !== "") {
                ssmSecret = data.SecretString;
                ssmSecret = JSON.parse(ssmSecret);
                ADDirectoryId = ssmSecret.AADDirectoryId;
                ClientId = ssmSecret.ClientId;
                RedirectUri = ssmSecret.RedirectUri;
                ClientSecret = ssmSecret.ClientSecret;
                PersonalAgentUserId = ssmSecret.PersonalAgentUserId;
                InvestmentAgentUserId = ssmSecret.InvestmentAgentUserId;
                var bookingAvailabilities = getO365Token(date, ADDirectoryId, ClientId, RedirectUri, ClientSecret, InvestmentAgentUserId, PersonalAgentUserId, callback, intentRequest, bookAppointment);
                return bookingAvailabilities;
            }
        }
    });

}

Array.prototype.remove = function () {
    var what, a = arguments, L = a.length, ax;
    while (L && this.length) {
        what = a[--L];
        while ((ax = this.indexOf(what)) !== -1) {
            this.splice(ax, 1);
        }
    }
    return this;
};

function BookAppointment(date, ADDirectoryId, ClientId, RedirectUri, ClientSecret, PersonalAgentUserId, InvestmentAgentUserId, callback, intentRequest, accesstoken) {
    var userId = "";
    if (appointmentType == "investment") {
        userId = InvestmentAgentUserId;
    }
    else {
        userId = PersonalAgentUserId;
    }

    var postUrl = "https://graph.microsoft.com/v1.0/users/" + userId + "/events";
    console.log("time : " + time);
    console.log("timeZone : " + timeZone);
    var endTime = parseInt(time) + 1;

    var pBody = JSON.stringify({
        "subject": "Customer meeting",
        "start": { "dateTime": date + "T" + time + ":00", "timeZone": timeZone },
        "end": { "dateTime": date + "T" + endTime + ":00:00", "timeZone": timeZone },
        "Attendees": [
            {
              "EmailAddress": {
                "Address": email
              },
              "Type": "Required"
            }
          ]
    });

    console.log("pBody : " + pBody);

    Request.post({
        "headers": {
            "Content-type": "application/json",
            "Authorization": "Bearer " + accesstoken
        },
        "url": postUrl,
        "body": pBody
    }, (error, response, postResBody) => {
        if (error) {
            return console.log(error);
        }

        console.log("Response Body " + postResBody);


        callback(close(outputSessionAttributes, 'Fulfilled', {
            contentType: 'PlainText',
            content: `Okay, I have booked your appointment.  We will see you at ${buildTimeOutputString(time)} on ${date}`
        }));

    });
}


function GetDateValues(date, ADDirectoryId, ClientId, RedirectUri, ClientSecret, PersonalAgentUserId, InvestmentAgentUserId, callback, intentRequest, accesstoken) {

    var userId = "";
    if (appointmentType == "investment") {
        userId = InvestmentAgentUserId;
    }
    else {
        userId = PersonalAgentUserId;
    }
    var getUrl = "https://graph.microsoft.com/v1.0/users/" + userId + "/calendarView?startDateTime=" + date + "T00:00:00.0000000&endDateTime=" + date + "T23:00:00.0000000";

    console.log("getUrl : " + getUrl);
    Request.get({
        "headers": { "Authorization": "Bearer " + accesstoken, "Prefer": " outlook.timezone=" + "\"" + timeZone + "\"" },
        "url": getUrl
    }, (error, response, getreqBody) => {
        if (error) {
            return console.log(error);
        }
        //console.log("Rahul : IN Completed 3");
        //console.log(getreqBody);
        //console.log("Value " + JSON.parse(getreqBody).value[0].start.dateTime.toString().split("T")[0]);
        //console.log("Value length" + JSON.parse(getreqBody).value.length);
        //console.log("Rahul : IN Completed 4");
        var calCount = 0;
        var standardSlots = ["10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00"];
        console.log("Calendar Values : " + getreqBody);
        while (calCount < JSON.parse(getreqBody).value.length) {
            var time = JSON.parse(getreqBody).value[calCount].start.dateTime.toString().split("T")[1].split(":")[0] + ":00";
            console.log("time " + calCount + " :" + time);
            standardSlots.remove(time);
            calCount = calCount + 1;
        }
        bookingMap[`${date}`] = standardSlots;

        if (standardSlots.length === 0) {
            //No availability on this day at all; ask for a new date and time.
            slots.Date = null;
            slots.Time = null;
            callback(elicitSlot(outputSessionAttributes, intentRequest.currentIntent.name, slots, 'Date',
                { contentType: 'PlainText', content: 'We do not have any availability on that date, is there another day which works for you?' },
                buildResponseCard('Specify Date', 'What day works best for you?',
                    buildOptions('Date', appointmentType, date, bookingMap))));
            return;
        }

        if (standardSlots.length === 1) {
            // If there is only one availability on the given date, try to confirm it.
            slots.Time = standardSlots[0];
            callback(confirmIntent(outputSessionAttributes, intentRequest.currentIntent.name, slots,
                { contentType: 'PlainText', content: `${messageContent}${buildTimeOutputString(standardSlots[0])} is our only availability, does that work for you?` },
                buildResponseCard('Confirm Appointment', `Is ${buildTimeOutputString(standardSlots[0])} on ${date} okay?`,
                    [{ text: 'yes', value: 'yes' }, { text: 'no', value: 'no' }])));
            return;
        }

        outputSessionAttributes.bookingMap = JSON.stringify(bookingMap);

        const availableTimeString = buildAvailableTimeString(standardSlots);
        callback(elicitSlot(outputSessionAttributes, intentRequest.currentIntent.name, slots, 'Time',
            { contentType: 'PlainText', content: `${messageContent}${availableTimeString}` },
            buildResponseCard('Specify Time', 'What time works best for you?',
                buildOptions('Time', appointmentType, availableTimeString, bookingMap))));
        return;

    });
}


function getO365Token(date, ADDirectoryId, ClientId, RedirectUri, ClientSecret, PersonalAgentUserId, InvestmentAgentUserId, callback, intentRequest, bookAppointment) {
    var reqBody = "client_id=" + ClientId + "&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&redirect_uri=" + RedirectUri + "&grant_type=client_credentials&client_secret=" + ClientSecret;
    var url = "https://login.microsoftonline.com/" + ADDirectoryId + "/oauth2/v2.0/token";

    //console.log("body " + reqBody);
    //console.log("url " + url);

    Request.post({
        "headers": { "content-type": "application/x-www-form-urlencoded" },
        "url": url,
        "body": reqBody,
    }, (error, response, body) => {
        if (error) {
            return console.log(error);
        }
        if (bookAppointment) {

            BookAppointment(date, ADDirectoryId, ClientId, RedirectUri, ClientSecret, PersonalAgentUserId, InvestmentAgentUserId, callback, intentRequest, JSON.parse(body).access_token);
        }
        else {
            GetDateValues(date, ADDirectoryId, ClientId, RedirectUri, ClientSecret, PersonalAgentUserId, InvestmentAgentUserId, callback, intentRequest, JSON.parse(body).access_token);
        }
    });


}

// Helper function to check if the given time and duration fits within a known set of availability windows.
// Duration is assumed to be one of 30, 60 (meaning minutes).  Availabilities is expected to contain entries of the format HH:MM.
function isAvailable(time, duration, availabilities) {
    if (duration === 30) {
        return (availabilities.indexOf(time) !== -1);
    } else if (duration === 60) {
        const secondHalfHourTime = incrementTimeByThirtyMins(time);
        return (availabilities.indexOf(time) !== -1 && availabilities.indexOf(secondHalfHourTime) !== -1);
    }
    // Invalid duration ; throw error.  We should not have reached this branch due to earlier validation.
    throw new Error(`Was not able to understand duration ${duration}`);
}

function isTimeAvailable(time, availabilities) {

    return (availabilities.indexOf(time) !== -1);
}


function getDuration(appointmentType) {
    const appointmentDurationMap = { personal: 60, investment: 60 };
    return appointmentDurationMap[appointmentType.toLowerCase()];
}

// Helper function to return the windows of availability of the given duration, when provided a set of 30 minute windows.
function getAvailabilitiesForDuration(duration, availabilities) {
    const durationAvailabilities = [];
    let startTime = '10:00';
    while (startTime !== '17:00') {
        if (availabilities.indexOf(startTime) !== -1) {
            if (duration === 30) {
                durationAvailabilities.push(startTime);
            } else if (availabilities.indexOf(incrementTimeByThirtyMins(startTime)) !== -1) {
                durationAvailabilities.push(startTime);
            }
        }
        startTime = incrementTimeByThirtyMins(startTime);
    }
    return durationAvailabilities;
}

function buildValidationResult(isValid, violatedSlot, messageContent) {
    return {
        isValid,
        violatedSlot,
        message: { contentType: 'PlainText', content: messageContent },
    };
}

function validateBookAppointment(appointmentType, date, time, email) {
    //console.log("Rahul 2 : " + appointmentType);

    if (appointmentType && !getDuration(appointmentType)) {
        return buildValidationResult(false, 'AppointmentType', 'I did not recognize that, can I book an appointment with a personal banking or an investment banking executive?');
    }
    if (time) {
        if (time.length !== 5) {
            return buildValidationResult(false, 'Time', 'I did not recognize that, what time would you like to book your appointment?');
        }
        const hour = parseInt(time.substring(0, 2), 10);
        const minute = parseInt(time.substring(3), 10);
        if (isNaN(hour) || isNaN(minute)) {
            return buildValidationResult(false, 'Time', 'I did not recognize that, what time would you like to book your appointment?');
        }
        if (hour < 10 || hour > 16) {
            // Outside of business hours
            return buildValidationResult(false, 'Time', 'Our business hours are ten a.m. to five p.m.  What time works best for you?');
        }
        if ([60, 0].indexOf(minute) === -1) {
            // Must be booked on the hour or half hour
            return buildValidationResult(false, 'Time', 'We schedule appointments every hour, what time works best for you?');
        }
    }
    if (date) {
        if (!isValidDate(date)) {
            return buildValidationResult(false, 'Date', 'I did not understand that, what date works best for you?');
        }
        if (parseLocalDate(date) <= new Date()) {
            return buildValidationResult(false, 'Date', 'Appointments must be scheduled a day in advance.  Can you try a different date?');
        } else if (parseLocalDate(date).getDay() === 0 || parseLocalDate(date).getDay() === 6) {
            return buildValidationResult(false, 'Date', 'Our office is not open on the weekends, can you provide a work day?');
        }
    }
    if(email == "")
    {
        return buildValidationResult(false, 'Email', 'Please specify your email address');
    }
    
    
    return buildValidationResult(true, null, null);
}

function buildTimeOutputString(time) {
    const hour = parseInt(time.substring(0, 2), 10);
    const minute = time.substring(3);
    if (hour > 12) {
        return `${hour - 12}:${minute} p.m.`;
    } else if (hour === 12) {
        return `12:${minute} p.m.`;
    } else if (hour === 0) {
        return `12:${minute} a.m.`;
    }
    return `${hour}:${minute} a.m.`;
}

// Build a string eliciting for a possible time slot among at least two availabilities.
function buildAvailableTimeString(availabilities) {
    let prefix = 'We have availabilities at ';
    if (availabilities.length > 3) {
        prefix = 'We have plenty of availability, including ';
    }
    prefix += buildTimeOutputString(availabilities[0]);
    if (availabilities.length === 2) {
        return `${prefix} and ${buildTimeOutputString(availabilities[1])}`;
    }
    return `${prefix}, ${buildTimeOutputString(availabilities[1])} and ${buildTimeOutputString(availabilities[2])}`;
}

// Build a list of potential options for a given slot, to be used in responseCard generation.
function buildOptions(slot, appointmentType, date, bookingMap) {
    const dayStrings = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    if (slot === 'AppointmentType') {
        return [
            { text: 'Personal', value: 'Personal' },
            { text: 'Investment', value: 'Investment' }
        ];
    } else if (slot === 'Date') {
        // Return the next five weekdays.
        const options = [];
        const potentialDate = new Date();
        while (options.length < 5) {
            potentialDate.setDate(potentialDate.getDate() + 1);
            if (potentialDate.getDay() > 0 && potentialDate.getDay() < 6) {
                options.push({
                    text: `${potentialDate.getMonth() + 1}-${potentialDate.getDate()} (${dayStrings[potentialDate.getDay()]})`,
                    value: potentialDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })
                });
            }
        }
        return options;
    } else if (slot === 'Time') {
        // Return the availabilities on the given date.
        if (!appointmentType || !date) {
            return null;
        }
        let availabilities = bookingMap[`${date}`];
        if (!availabilities) {
            return null;
        }
        availabilities = getAvailabilitiesForDuration(getDuration(appointmentType), availabilities);
        if (availabilities.length === 0) {
            return null;
        }
        const options = [];
        for (let i = 0; i < Math.min(availabilities.length, 5); i++) {
            options.push({ text: buildTimeOutputString(availabilities[i]), value: buildTimeOutputString(availabilities[i]) });
        }
        return options;
    }
}

// --------------- Functions that control the skill's behavior -----------------------

/**
 * Performs dialog management and fulfillment for booking a dentists appointment.
 *
 * Beyond fulfillment, the implementation for this intent demonstrates the following:
 *   1) Use of elicitSlot in slot validation and re-prompting
 *   2) Use of confirmIntent to support the confirmation of inferred slot values, when confirmation is required
 *      on the bot model and the inferred slot values fully specify the intent.
 */

var bookingMap = "";
var outputSessionAttributes;
var slots = "";
var messageContent = "";
var appointmentType = "";
var date;
var time;
var email;

function makeAppointment(intentRequest, callback) {
    appointmentType = intentRequest.currentIntent.slots.AppointmentType;
    //console.log("Rahul " + appointmentType);
    date = intentRequest.currentIntent.slots.Date;
    time = intentRequest.currentIntent.slots.Time;
    email = intentRequest.currentIntent.slots.Email;
    const source = intentRequest.invocationSource;
    outputSessionAttributes = intentRequest.sessionAttributes || {};
    bookingMap = JSON.parse(outputSessionAttributes.bookingMap || '{}');
    //console.log("10");
    if (source === 'DialogCodeHook') {
        // Perform basic validation on the supplied input slots.
        slots = intentRequest.currentIntent.slots;
        const validationResult = validateBookAppointment(appointmentType, date, time, email);
        if (!validationResult.isValid) {
            slots[`${validationResult.violatedSlot}`] = null;
            callback(elicitSlot(outputSessionAttributes, intentRequest.currentIntent.name,
                slots, validationResult.violatedSlot, validationResult.message,
                buildResponseCard(`Specify ${validationResult.violatedSlot}`, validationResult.message.content,
                    buildOptions(validationResult.violatedSlot, appointmentType, date, bookingMap))));
            return;
        }

        if (!appointmentType) {
            callback(elicitSlot(outputSessionAttributes, intentRequest.currentIntent.name,
                intentRequest.currentIntent.slots, 'AppointmentType',
                { contentType: 'PlainText', content: 'Would you like to meet a personal banking executive or an investment one?' },
                buildResponseCard('Specify Appointment Type', 'Which type of bank executive would you like to meet?',
                    buildOptions('AppointmentType', appointmentType, date, null))));
            return;
        }
        if (appointmentType && !date) {
            callback(elicitSlot(outputSessionAttributes, intentRequest.currentIntent.name,
                intentRequest.currentIntent.slots, 'Date',
                { contentType: 'PlainText', content: `When would you like to meet your ${appointmentType} banking executive?` },
                buildResponseCard('Specify Date', `When would you like to meet your ${appointmentType} banking executive?`,
                    buildOptions('Date', appointmentType, date, null))));
            return;
        }

        if (appointmentType && date) {

            let bookingAvailabilities = bookingMap[`${date}`];
            if (bookingAvailabilities == null) {
                getO365APISecretsAndOrchestrate(date, callback, intentRequest, false);
                return;
            }
            if (time) {

                outputSessionAttributes.formattedTime = buildTimeOutputString(time);
                // Validate that proposed time for the appointment can be booked by first fetching the availabilities for the given day.  To
                // give consistent behavior in the sample, this is stored in sessionAttributes after the first lookup.
                if (isTimeAvailable(time, bookingAvailabilities)) {
                    callback(delegate(outputSessionAttributes, slots));
                    return;
                }
                else {
                    messageContent = 'The time you requested is not available. ';
                    slots.Time = null;
                    callback(elicitSlot(outputSessionAttributes, intentRequest.currentIntent.name, slots, 'Time',
                        { contentType: 'PlainText', content: messageContent },
                        buildResponseCard('Specify Time', 'What other time works for you?',
                            buildOptions('Time', appointmentType, time, bookingMap))));
                    return;
                }
            }
        }
        callback(delegate(outputSessionAttributes, slots));
        return;
    }

    // Book the appointment.  In a real bot, this would likely involve a call to a backend service.
    if (appointmentType && date && time) {
        getO365APISecretsAndOrchestrate(date, callback, intentRequest, true);
    }
}

// --------------- Intents -----------------------

/**
 * Called when the user specifies an intent for this skill.
 */
function dispatch(intentRequest, callback) {
    // console.log(JSON.stringify(intentRequest, null, 2));
    //console.log(`dispatch userId=${intentRequest.userId}, intent=${intentRequest.currentIntent.name}`);

    const name = intentRequest.currentIntent.name;

    // Dispatch to your skill's intent handlers
    if (name === 'MakeAppointment') {
        return makeAppointment(intentRequest, callback);
    }
    throw new Error(`Intent with name ${name} not supported`);
}

// --------------- Main handler -----------------------

function loggingCallback(response, originalCallback) {
    originalCallback(null, response);
}

var timeZone = "";
var timeZoneOutlook = ";"
// Route the incoming request based on intent.
// The JSON body of the request is provided in the event slot.
exports.handler = (event, context, callback) => {
    try {
        // By default, treat the user request as coming from the America/New_York time zone.
        process.env.TZ = 'America/New_York';
        timeZone = process.env.Time_Zone;
        o365APISecrets = process.env.o365_API_Secrets
        timeZoneOutlook = process.env.Time_Zone_Outlook;
        dispatch(event, (response) => loggingCallback(response, callback));
    } catch (err) {
        console.log("Error :" + err);
        callback(err);
    }
};

