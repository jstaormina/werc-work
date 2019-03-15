const Excel = require('exceljs');
const moment = require('moment');
const readline = require('readline');
const google = require('googleapis').google;
const fs = require('fs');

const SCOPES = ['https://www.googleapis.com/auth/calendar.events'];
const TOKEN_PATH = 'token.json';

const colorMap = {
    'FFBFBFBF': 'HR1',
    'FFC0C0C0': 'HR1',
    'FF00FFFF': 'HR2',
    'FFFFFF00': 'HR4',
    'FFFF99CC': 'HR5a',
    'FFFF0000': 'HR5b',
    'FF00FF00': 'HR various',

};

(async () => {
    const events = await parseXlsx();
    await updateCalendarWithEvent(events);
})();

async function parseXlsx() {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile('./resources/ATC_RUN_SCHEDULES.xlsx');
    const worksheet = workbook.getWorksheet('Full_Marathon');
    const calEntries = [];
    worksheet.eachRow({}, (row, rowNumber) => {
        if (rowNumber > 8 && rowNumber < 28) {
            const startDate = row.values[1].result;
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                if (colNumber > 3 && colNumber < 11 && cell.value !== 'OFF') {
                    const calEntry = {
                        summary: 'Marathon Training',
                        start: {
                            dateTime: moment(startDate).add(colNumber - 3, 'days').hours(7).toDate()
                        },
                        end: {
                            dateTime: moment(startDate).add(colNumber - 3, 'days').hours(8).toDate()
                        }
                    };
                    const hr = cell.style.fill ? colorMap[cell.style.fill.fgColor.argb] : 'HR3';
                    if (colNumber === 4 || colNumber === 9) {
                        calEntry.description = 'Run ' + cell.value + ' miles at ' + hr;
                    } else if (colNumber === 5) {
                        calEntry.description = 'Run ' + cell.value + ' minutes at ' + hr;
                    } else if (colNumber === 6) {
                        calEntry.description = 'Bike ' + cell.value + ' hour at ' + hr;
                    } else if (colNumber === 7 && cell.value !== 1) {
                        calEntry.description = 'low intensity recovery workout ' + cell.value + ' minutes at ' + hr;
                    } else if (colNumber === 7) {
                        calEntry.description = 'Bootcamp or track ' + cell.value + ' hour at ' + hr;
                    } else if (cell.value === 'XT') {
                        calEntry.description = 'Crosstrain';
                    }

                    calEntries.push(calEntry);
                }
            });
        }
    });

    return calEntries;
}

function updateCalendarWithEvent(events) {
    fs.readFile('credentials.json', (err, content) => {
        if (err) return console.log('Error loading client secret file:', err);
        authorize(JSON.parse(content), async (auth) => {
            await removeAllTrainingEvents(auth, events[0].start.dateTime, events[events.length - 1].end.dateTime);
            await createEvents(auth, events);
        });
    });
}

function authorize(credentials, callback) {
    const {client_secret, client_id, redirect_uris} = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]);

    fs.readFile(TOKEN_PATH, (err, token) => {
        if (err) return getAccessToken(oAuth2Client, callback);
        oAuth2Client.setCredentials(JSON.parse(token));
        callback(oAuth2Client);
    });
}

function getAccessToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error retrieving access token', err);
            oAuth2Client.setCredentials(token);
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

async function createEvents(auth, workoutEvents) {
    const calendar = google.calendar({version: 'v3', auth});
    for (const event of workoutEvents) {
        await timeout(200);
        try {
            const response = await new Promise((resolve, reject) => {
                calendar.events.insert({
                    'calendarId': 'primary',
                    'resource': event
                }, (err, response) => {
                    if (err) {
                        reject(err);
                    }
                    resolve(response);
                });
            });
            console.log('New event created: ' + response.data.htmlLink);
        } catch (err) {
            console.log(err.message);
        }
    }
}

async function removeAllTrainingEvents(auth, timeMin, timeMax) {
    try {
        const calendar = google.calendar({version: 'v3', auth});
        const workoutEvents = await getExistingEvents(auth, timeMin, timeMax);
        for (const event of workoutEvents) {
            await timeout(200);
            try {
                await new Promise((resolve, reject) => {
                    calendar.events.delete({
                        'calendarId': 'primary',
                        'eventId': event.id
                    }, (err, response) => {
                        if (err) {
                            reject(err);
                        }
                        resolve(response);
                    });
                });
                console.log('Event Removed: ' + event.id);
            } catch (err) {
                console.log(err.message);
            }
        }
    } catch (err) {
        console.log('The API returned an error: ' + err);
    }
}

function getExistingEvents(auth, timeMin, timeMax) {
    const calendar = google.calendar({version: 'v3', auth});
    return new Promise((resolve, reject) => {
        calendar.events.list({
            calendarId: 'primary',
            q: 'Marathon Training',
            timeMin: timeMin.toISOString(),
            timeMax: timeMax.toISOString(),
            maxResults: 200,
            singleEvents: true,
            orderBy: 'startTime',
        }, (err, res) => {
            if (err) {
                reject(err);
            }
            const events = res.data.items;
            resolve(events);
        });
    });
}

function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}