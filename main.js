const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const https = require("https");
const querystring = require("querystring");

require("dotenv").config({ path: __dirname + "/.env" });

var time_interval = parseInt(process.env.time_interval); // polling time interval in millis

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
		path: '/' + process.env.token_domain + '/oauth2/v2.0/token',
		method: 'POST',
		headers: {
			'Content-Type': 'application/x-www-form-urlencoded',
			'Content-Length': data.length
		}
	}

	const req = https.request(options, (res) => {
		res.on('data', (d) => {
			resp = JSON.parse(d.toString('utf8'));
			callback({
				token: resp.access_token,
				expiry: resp.expires_in
			}
			);
		})
	})

	req.on('error', (error) => {
		console.error(error)
	});

	req.end(data);

}

function patchAlert(alert) {
	client.api('/security/alerts/' + alert.id).patch({
		"assignedTo": "SyslogForwarder",
		"closedDateTime": new Date(Date.now()).toISOString(),
		"comments": alert.comments,
		"tags": alert.tags,
		"feedback": "unknown",
		"status": "inProgress",
		"vendorInformation": {
			"provider": alert.vendorInformation.provider,
			"providerVersion": alert.vendorInformation.providerVersion,
			"subProvider": alert.vendorInformation.subProvider,
			"vendor": alert.vendorInformation.vendor
		}
	},
		(err, res) => {
			console.log(err); // prints info about authenticated user
			console.log(res); // prints info about authenticated user
			// callback({err, res});
		});
}

//getting the (new) token
async function refresh() {
	return new Promise(function (resolve, reject) {
		getToken((c) => {
			token = c.token;
			expiry = parseInt(c.expiry);

			client = MicrosoftGraph.Client.init({
				authProvider: (done) => {
					done(null, token); //first parameter takes an error if you can't get an access token
				}
			});
			setTimeout(function () { refresh() }, (expiry * 1000)); //try to get a new token after <expiry>ms * 1000 ms
			resolve(token)
		});
	});
};

async function getAlertsAPI(top, skip) {
	return new Promise((resolve, reject) => {
		client.api("/security/alerts").filter("status eq 'newAlert'").top(top).skip(skip).get()
		.then(res => {
			resolve(res);
		})
		.catch((err) => {
			reject(err);
		});
	});
}


function sendAndPatchAlerts(securityAlerts){
	try{
		for (var i = 0; i <securityAlerts.value.length; i++){
			console.log(securityAlerts.value[i]);
		    syslogSend(securityAlerts.value[i]);
		    patchAlert(securityAlerts.value[i]);
		}
	}catch(err){
		console.log("Exception occured when accessing or sending alert data: " + err);
		console.log("Is securityAlerts object still what you expected?");
	}
	
}


async function getAlerts() {

	let moreAlerts = false;
	let _top = 5; //number of alerts to pull at a time
	let _skip = 0; //number of alerts to skip (offset) default to 0

	do{
		let securityAlerts = await getAlertsAPI(_top,_skip);
	
		console.log("in getalerts");
		console.log(securityAlerts);
	
		//sendAndPatchAlerts(securityAlerts);
	
		//Check if there are more alerts by checking the 'nextLink' in the returned obj
		//if nextLink is null = no more alerts 
		let nextLink = securityAlerts["@odata.nextLink"];
		if(nextLink != null){
			//Extract top and skip values from the URL
			//example: https://graph.microsoft.com/v1.0/security/alerts?$filter=status+eq+%27newAlert%27&$top=5&$skip=5 (the 5 and 5)
			_top = parseInt(nextLink.split('&')[1].split('=')[1]);
			_skip = parseInt(nextLink.split('&')[2].split('=')[1]);
			moreAlerts = true;
			console.log("fetching more.. top="+_top+" skip="+_skip);
		}
		else{
			console.log("no more alerts!");
			moreAlerts = false;
		}

	}while(moreAlerts);
}

function syslogSend(alertMessage) {
	// Initialising syslog
	var syslog = require("syslog-client");

	// Getting environment variables
	var SYSLOG_SERVER = SYSLOG_SERVER;
	var SYSLOG_PROTOCOL = SYSLOG_PROTOCOL;
	var SYSLOG_HOSTNAME = SYSLOG_HOSTNAME;
	var SYSLOG_PORT = SYSLOG_PORT;

	// Options for syslog connection
	var options = {
		syslogHostname: SYSLOG_HOSTNAME,
		transport: SYSLOG_PROTOCOL,
		port: SYSLOG_PORT
	};

	// Create syslog client
	var client = syslog.createClient(SYSLOG_SERVER, options);

	// Send syslog message
	client.log(JSON.stringify(alertMessage), options, function (error) {
		if (error) {
			console.log(error);
		} else {
			console.log("Sent message successfully");
		}
	});
};

async function main() {
	if (token == null || expiry == null || expiry < Date.now) {
		await refresh();
		getAlerts();
	}

	setInterval(function () {
		getAlerts();
	}, time_interval);
}

main();
