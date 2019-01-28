const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const https = require("https");
const querystring = require("querystring");

require('dotenv').config({path: __dirname + '/.env'});

var time_interval = 60000; // polling time interval in millis

var client, token, expiry;

function getToken(callback) {

    var resp;

    var data = querystring.stringify({
        tenant: process.env.tenant,
        client_id: process.env.client_id,
        scope: process.env.scope,
        client_secret: process.env.client_secret,
        grant_type: process.env.grant_type
    });
    
    var options = {
        host: 'login.microsoftonline.com',
        port: '443',
        path: '/b84melive.onmicrosoft.com/oauth2/v2.0/token',
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': data.length
        }
    };
  
    const req = https.request(options, (res) => {
        res.on('data', (d) => {
            resp = JSON.parse(d.toString('utf8'));
            callback({
                token: resp.access_token,
                expiry: Date.now() + parseInt(resp.expires_in) * 1000}
            );
        })
    })
    
    req.on('error', (error) => {
        console.error(error)
    })
    
    req.end(data);
}

async function refresh() {
    return new Promise(function(resolve, reject) {
        getToken((c) => {
            token = c.token;
            expiry = parseInt(c.expiry);
    
            client = MicrosoftGraph.Client.init({
                authProvider: (done) => {
                    done(null, token); //first parameter takes an error if you can't get an access token
                }
            });
    
            setTimeout(function(){ refresh() }, ( Date.now + parseInt(c.expiry) ));
            resolve(token)
        });      
    });
};

function getAlerts() {
    client.api('/security/alerts').top(1).get((err, res) => {
        console.log(err); // prints info about authenticated user
        console.log(res); // prints info about authenticated user
    });
}

async function main() {
    if (token == null || expiry == null || expiry < Date.now){
        await refresh();
        getAlerts();
    }
    
    setInterval(function() {
        getAlerts();
    }, time_interval);
}

main();