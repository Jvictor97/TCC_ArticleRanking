const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const xl = require('exceljs');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Google Sheets API.
  authorize(JSON.parse(content), recursiveQuestion);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);

  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

var recursiveQuestion = function(auth) {
  let ans = '';
  rl.question('Calculate relevance? (y|n) ', function(answer) {
    ans = answer;
    
    if(ans === 'n'){
      createGoogleSheet();
    }
    else if (ans === 'y') {
      defineFonteArtigos(auth);
    }
    else {
      console.log("Wrong Input!!");
      recursiveQuestion();
    }
  });
}

var defineFonteArtigos = function (auth){
  rl.question('Excel (x) or Google Sheets (g) ? ', function(answer) {
    if(answer == 'x'){
      abreExcel();
      return;
    }
    else if (answer == 'g') {
      calculaRankingGoogleSheets(auth);
    }
    else {
      console.log("Wrong Input!!");
      defineFonteArtigos();
    }
  });
}

function abreExcel(){
  rl.question('Type in the xlsx filename: ', (answer) => {
    rl.close();
    calculaRankingExcel(answer);
    return;
  });
}

function calculaRankingExcel(filename) {
  // var path = require('path');

  // let options = {
  //   filename: path.resolve(filename),
  // };

  // let wb = new xl.stream.xlsx.WorkbookWriter(options);

  let wb = new xl.Workbook();
  var path = require('path');
  var filePath = path.resolve(filename);
  
  console.log(`Opening file '${filename}'...`)

  wb.xlsx.readFile(filePath)
  
  .then( function() {
        console.log('Calculating Ranking...')
        let ws = wb.getWorksheet(1);
        let maiorFator = 1;
        let maiorVel = 1;
        let total = 0; 
        let seen = 0;

        ws.eachRow( (row, idx) => {
            console.log("Estou na linha: "+idx);
            // let fator = isNaN(row.getCell('F')) ? 0 : row.getCell('F');
            // let vel = isNaN(row.getCell('G')) ? 0 : row.getCell('G');

            // total = isNaN(row.getCell('F')) ? total : total+1;
        
            // maiorFator = fator > maiorFator ? fator : maiorFator;
            // maiorVel = vel > maiorVel ? vel : maiorVel;
        });

        //ws.eachRow( (row, idx) => {
            // let ano = isNaN(row.getCell('E')) ? 0 : row.getCell('E') / 2019;
            // let fator = isNaN(row.getCell('F')) ? 0 : row.getCell('F') / maiorFator;
            // let vel = isNaN(row.getCell('G')) ? 0 : row.getCell('G') / maiorVel;

            // let result = ano + fator + vel;
            
            // row.getCell('J').value = result;

            // seen++;

            // process.stdout.write(`Total: ${seen}/${total} (${seen/total*100}%)       \r`)
        //});

        wb.xlsx.writeFile('ranking.xlsx')
          .then( () => console.log('*** Ranking Finished ***') )
    });
}


/**
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
function calculaRankingGoogleSheets(auth) {

  console.log("Processing ranking...");

  let delay = 0;
  const sheets = google.sheets({version: 'v4', auth});
  sheets.spreadsheets.values.get({
    spreadsheetId: '17-ohbIToQjbOFY5MIgH_aP2XVIqOGhKM5O8nYufQiok',
    range: 'Ranking!E2:G10',
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);
    const rows = res.data.values;
    if (rows.length) {
      let maiorFator = 1;
      let maiorVel = 1;
      let total = 0;
      let seen = 0;
      
      rows.map((row) => {
        let fator = isNaN(row[1]) ? 0 : row[1];
        let vel = isNaN(row[2]) ? 0 : row[2];

        total = isNaN(row[1]) ? total : total+1;
        
        maiorFator = fator > maiorFator ? fator : maiorFator;
        maiorVel = vel > maiorVel ? vel : maiorVel;
      })
      
      rows.map((row, idx) => {
        if(row[0] && row[1] && row[2]) {
          let ano = isNaN(row[0]) ? 0 : row[0] / 2019;
          let fator = isNaN(row[1]) ? 0 : row[1] / maiorFator;
          let vel = isNaN(row[2]) ? 0 : row[2] / maiorVel;

          let result = ano + fator + vel;

          setTimeout( () => {
            sheets.spreadsheets.values.update({
              spreadsheetId: '17-ohbIToQjbOFY5MIgH_aP2XVIqOGhKM5O8nYufQiok',
              range: `Ranking!I${idx+2}:I${idx+2}`,
              valueInputOption: 'RAW',
              resource: {
                values: [ 
                 [ result ], 
                ]
              },
              auth: auth
            }, (err, res) => {
              if(err) {
                console.log('ERRO')
                //process.exit(1);
              }
              seen++;
              let date = new Date(null);
              date.setSeconds((total-seen)*2);
              var time = date.toISOString().substr(11, 8);
              process.stdout.write(`Total: ${seen}/${total} (${seen/total*100}%) - ${time} \r`);
            })
          }, delay);
          delay += 2000;
        }
      });
    } else {
      console.log('No data found.');
    }
  });

  createGoogleSheet();
}

function createGoogleSheet(){
  console.log("CreateSheet...")
}