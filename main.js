const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const https = require("https");
const querystring = require("querystring");

require("dotenv").config({ path: __dirname + "/.env" });

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

  var options = {
    host: "login.microsoftonline.com",
    port: "443",
    path: "/b84melive.onmicrosoft.com/oauth2/v2.0/token",
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      "Content-Length": data.length
    }
  };

  console.log("1");
  const req = https.request(options, res => {
    console.log(`statusCode: ${res.statusCode}`);

    res.on("data", d => {
      process.stdout.write(d);
    });
  });

  console.log("2");
  // req.on('error', (error) => {
  // console.error(error)
  // })

  // req.write()

  console.log("3");
  req.end(data);
  console.log("4");
}

function syslogSend(context, myEventHubMessage) {
  // Initialising syslog
  var syslog = require("syslog-client");

  // Getting environment variables
  var SYSLOG_SERVER = GetEnvironmentVariable("SYSLOG_SERVER");
  var SYSLOG_PROTOCOL = syslog.Transport.Udp;
  var SYSLOG_HOSTNAME = GetEnvironmentVariable("SYSLOG_HOSTNAME");
  var SYSLOG_PORT = GetEnvironmentVariable("SYSLOG_PORT");

  // Options for syslog connection
  var options = {
    syslogHostname: SYSLOG_HOSTNAME,
    transport: SYSLOG_PROTOCOL,
    port: SYSLOG_PORT
  };

  // Log connection variables
  context.log("SYSLOG Server: ", SYSLOG_SERVER);
  context.log("SYSLOG Port: ", SYSLOG_PORT);
  context.log("SYSLOG Protocol: ", SYSLOG_PROTOCOL);
  context.log("SYSLOG Hostname: ", SYSLOG_HOSTNAME);

  // log received message from event hub
  context.log("Event Hubs trigger function processed message: ",myEventHubMessage);

  // cycle through eventhub messages and send syslog
  for (var i = 0; i < myEventHubMessage.records.length; i++) {
    var l = myEventHubMessage.records[i];
    client.log(JSON.stringify(l), options, function(error) {
      if (error) {
        context.log(error);
      } else {
        context.log("Sent message successfully");
      }
    });
  }
};

// Request alerts from the Microsoft Graph API
function graphCall(){
  var options = {
    host: "graph.microsoft.com",
    port: "443",
    path: "/v1.0/security/alerts?$top=1000",
    method: "GET",
    headers: {
      "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFDRWZleFh4amFtUWIzT2VHUTRHdWd2YlZJUXRlSi1MVFFuTmFlT2ltcEplYTBCNmZlOXFrc2s4UGxSMG5jU0JiSUo4aG5FNzhGV3hnQ2NuazJMOHE2WVp4dmdwM21fcTl4amVZRzBsQjd0a0NBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoibmJDd1cxMXczWGtCLXhVYVh3S1JTTGpNSEdRIiwia2lkIjoibmJDd1cxMXczWGtCLXhVYVh3S1JTTGpNSEdRIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC81YTBjN2ZjNy01NmFkLTQxOTctYTdmYS0xNjc5ZWMwNDA1ZjAvIiwiaWF0IjoxNTQ4NzA5MzYwLCJuYmYiOjE1NDg3MDkzNjAsImV4cCI6MTU0ODcxMzI2MCwiYWlvIjoiNDJKZ1lLai8zL3J6NWNRYmdhdC9QYWp6WGFEUUR3QT0iLCJhcHBfZGlzcGxheW5hbWUiOiJtc2cyc3lzbG9nIiwiYXBwaWQiOiI3Y2U2ZjcyMS03OTI3LTRmM2YtYjMwZS0xZDE0NzAxOTMxNTYiLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC81YTBjN2ZjNy01NmFkLTQxOTctYTdmYS0xNjc5ZWMwNDA1ZjAvIiwib2lkIjoiNWZlYTMwODItMTM5OC00YzFkLTg5ZjItNWRmNWU4ODkwMDdlIiwicm9sZXMiOlsiU2VjdXJpdHlFdmVudHMuUmVhZC5BbGwiLCJTZWN1cml0eUV2ZW50cy5SZWFkV3JpdGUuQWxsIl0sInN1YiI6IjVmZWEzMDgyLTEzOTgtNGMxZC04OWYyLTVkZjVlODg5MDA3ZSIsInRpZCI6IjVhMGM3ZmM3LTU2YWQtNDE5Ny1hN2ZhLTE2NzllYzA0MDVmMCIsInV0aSI6IjAzWVBWRmRJdGtpNDZYaEhIRXFCQUEiLCJ2ZXIiOiIxLjAiLCJ4bXNfdGNkdCI6MTUzMTQ4NTU4OH0.SqZj3_cRBFXavl-IHefetgnmGthwBmRmWNA7G--2JJg0W5-bQyG4O-0Umkz-E2gxEN0cOh60XbJbSF5CKOD-24nvNymfSw5oo9pmMPmH0BbupCN2HY9Hw3PnoWf3WXd6vq1vVg7jjL7YL1mVeXulKIawh1iqjL5voHfwYzSSBTMr6d3ZRYPWUpZvRIH8Vx9AKryvjAkQsr8vZgCstrmvtBfz5ZzaaIN7eiQIvAx5Axp9wjkIj2R06k-BLR5UQt8WbehGMjMeSuBjjLnG8bBbl_Lc1_qUSAmIn3NgcOw0r4_1yETdkMtNqbyCSHjWaz0uUnNtJxnSTqzcZ3yCYS2-gw"
    }
	};

	const graphReq = https.request(options, (res) => {
		console.log('statusCode:', res.statusCode);

		res.on('data', (d) => {
			process.stdout.write(d);
		});
	});
	
	graphReq.on('error', (e) => {
		console.error(e);
	});
	graphReq.end();

}

//getToken();
graphCall();

var client = MicrosoftGraph.Client.init({
  authProvider: done => {
    done(null, "PassInAccessTokenHere"); //first parameter takes an error if you can't get an access token
  }
});

client
  .api("/security/alerts")
  .top(1000)
  .get((err, res) => {
    console.log(res); // prints info about authenticated user
  });
