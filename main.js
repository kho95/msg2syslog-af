const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const https = require("https");
const querystring = require("querystring");

require('dotenv').config({path: __dirname + '/.env'});

var token, expiry;

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

getToken((c) => {
    token = c.token;
    expiry = c.expiry;

    var client = MicrosoftGraph.Client.init({
        authProvider: (done) => {
            done(null, token); //first parameter takes an error if you can't get an access token
        }
    });
    
    client.api('/security/alerts').top(1000).get((err, res) => {
        console.log(err); // prints info about authenticated user
        console.log(res); // prints info about authenticated user
    });
});



