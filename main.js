const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const https = require("https");
const querystring = require("querystring");

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
        tenant: '5a0c7fc7-56ad-4197-a7fa-1679ec0405f0',
        client_id: '7ce6f721-7927-4f3f-b30e-1d1470193156',
        scope: 'https://graph.microsoft.com/.default',
        client_secret: 'jhPOEH71lqmkfBBF433%=-+',
        grant_type: 'client_credentials'
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