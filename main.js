const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const https = require("https");
const querystring = require("querystring");
require("dotenv").config({ path: __dirname + "/.env" });

var time_interval = parseInt(process.env.time_interval); // polling time interval in millis
var client, token, expiry;

/* Get Microsoft Graph Authentication Token */
function getToken(callback) {
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
			var resp = JSON.parse(d.toString('utf8'));
			callback({
				token: resp.access_token,
				expiry: resp.expires_in
			});
		})
	})

	req.on('error', (error) => {
        console.log(error);
        var timeout = 30000;
        console.log("Trying again in " + timeout + " milliseconds");
        req.end();
		setTimeout(getToken, timeout);
	});

	req.end(data);
}

/* Mark alert on Microsoft Graph Security as forwarded */
function patchAlert(alert) {
    console.log("Patching " + alert.id);
	client.api('/security/alerts/' + alert.id).patch({
		"assignedTo": "SyslogForwarder",
		"closedDateTime": new Date(Date.now()).toISOString(),
		"comments": alert.comments,
		"tags": alert.tags,
		"feedback": "unknown",
		"status": "newAlert",
		"vendorInformation": {
			"provider": alert.vendorInformation.provider,
			"providerVersion": alert.vendorInformation.providerVersion,
			"subProvider": alert.vendorInformation.subProvider,
			"vendor": alert.vendorInformation.vendor
		}
	},
    (err, res) => { });
}

/* Get the (new) authentication token */
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

/* Get new alerts by 'page' */
async function getAlertsAPI(top, skip) {
	return new Promise((resolve, reject) => {
		client.api("/security/alerts").filter("status eq 'inProgress'").top(top).skip(skip).get()
		.then(res => {
			resolve(res);
		})
		.catch((err) => {
			reject(err);
		});
	});
}

/* Send alerts to Syslog and patch alerts on Microsoft Graph */
function sendAndPatchAlerts(securityAlerts){
	try {
		for (var i = 0; i <securityAlerts.value.length; i++) {
		    syslogSend(securityAlerts.value[i]);
		    patchAlert(securityAlerts.value[i]);
		}
	} catch (err) {
		console.log("Exception occured when accessing or sending alert data: " + err);
	}
}

/* Get all new alerts */
async function getAlerts() {

	let moreAlerts = false;
	let _top = 1000; // number of alerts to pull at a time (max 1000)
	let _skip = 0; // number of alerts to skip (offset) default to 0

	do {
		let securityAlerts = await getAlertsAPI(_top,_skip);
		console.log(securityAlerts);
		sendAndPatchAlerts(securityAlerts);
	
		//Check if there are more alerts by checking the 'nextLink' in the returned obj
		//if nextLink is null = no more alerts 
		let nextLink = securityAlerts["@odata.nextLink"];
		if (nextLink != null) {
			//Extract top and skip values from the URL
			//example: https://graph.microsoft.com/v1.0/security/alerts?$filter=status+eq+%27newAlert%27&$top=5&$skip=5 (the 5 and 5)
			_top = parseInt(nextLink.split('&')[1].split('=')[1]);
			_skip = parseInt(nextLink.split('&')[2].split('=')[1]);
			moreAlerts = true;
			console.log("fetching next top="+_top+" skip="+_skip);
		}
		else {
			console.log("no more alerts!");
			moreAlerts = false;
		}
	} while (moreAlerts);
}

function syslogSend(alert) {
	console.log("Sending Syslog " + alert.id);
	// Initialising syslog
	var syslog = require("syslog-client");

	// Getting environment variables
	var SYSLOG_SERVER = SYSLOG_SERVER;

	// Options for syslog connection
	var options = {
		syslogHostname: "SyslogForwarder",
		transport: "UDP",
		port: 514
	};

	// Create syslog client
	var client = syslog.createClient(SYSLOG_SERVER, options);

	// Send syslog message
	client.log(JSON.stringify(alert), options, function (error) {
		if (error) {
			console.log(error);
		} else {
			console.log("Sent " + alert.id + " successfully!");
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
