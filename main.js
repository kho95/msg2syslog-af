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

function getToken() {

    var data = querystring.stringify({
        tenant: process.env.tenant,
        client_id: process.env.client_id,
        scope: process.env.scope,
        client_secret: process.env.client_secret,
        grant_type: process.env.grant_type
    });

    console.log(process.env.tenant)
    console.log(process.env.client_id)
    console.log(process.env.scope)
    
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
  
    console.log("1")
    const req = https.request(options, (res) => {
        console.log(`statusCode: ${res.statusCode}`)
        
        res.on('data', (d) => {
            process.stdout.write(d)
        })
    })
    
    console.log("2")
    // req.on('error', (error) => {
    // console.error(error)
    // })
    
    req.write(data)

    console.log("3")
    req.end()
    console.log("4")
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