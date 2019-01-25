const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const https = require("https");
const querystring = require("querystring");

require('dotenv').config({path: __dirname + '/.env'});

// const http = require('http');
// const hostname = '127.0.0.1';
// const port = 3000;
// const server = http.createServer((req, res) => {
//   res.statusCode = 200;
//   res.setHeader('Content-Type', 'text/plain');
//   res.end('Hello World\n');
// });
// server.listen(port, hostname, () => {
//   console.log(`Server running at http://${hostname}:${port}/`);
// });

var token;
var expiry;

function getToken() {

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
        // console.log(`statusCode: ${res.statusCode}`)
        res.on('data', (d) => {
            var resp = JSON.parse(d.toString('utf8'));
            token = resp.access_token;
            expiry = resp.expires_in;
            console.log(token);
        })
    })
    
    req.on('error', (error) => {
        console.error(error)
    })
    
    req.end(data)
}

getToken();

var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
        done(null, "PassInAccessTokenHere"); //first parameter takes an error if you can't get an access token
    }
});

client.api('/security/alerts').top(1000).get((err, res) => {
    console.log(res); // prints info about authenticated user
});