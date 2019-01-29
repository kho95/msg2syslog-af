const MicrosoftGraph = require("@microsoft/microsoft-graph-client");
const https = require("https");
const querystring = require("querystring");
require("dotenv").config({ path: __dirname + "/.env" });

var time_interval = parseInt( GetEnvironmentVariable(TIME_INTERVAL) ); // polling time interval in millis
var client, token, expiry;

/* Get Microsoft Graph Authentication Token */
function getToken(callback) {
	var data = querystring.stringify({
		tenant:  GetEnvironmentVariable(MSG_TENANT),
		client_id: GetEnvironmentVariable(MSG_CLIENT_ID),
		scope: GetEnvironmentVariable(MSG_SCOPE) || "https://graph.microsoft.com/.default",
		client_secret: GetEnvironmentVariable(MSG_CLIENT_SECRET),
		grant_type: GetEnvironmentVariable(MSG_GRANT_TYPE) || client_credentials
	});

	var options = {
		host: 'login.microsoftonline.com',
		port: '443',
		path: '/' + GetEnvironmentVariable(TOKEN_DOMAIN) + '/oauth2/v2.0/token',
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
		setTimeout(getToken(() => {}), timeout);
	});

	req.end(data);
}

/* Mark alert on Microsoft Graph Security as forwarded */
async function patchAlert(alert) {
	return new Promise(function (resolve, reject) {
		console.log("Patching " + alert.id);
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
            resolve(res);
        });
	});
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
		client.api("/security/alerts").filter("status eq 'newAlert'").top(top).skip(skip).get()
		.then(res => {
			resolve(res);
		})
		.catch((err) => {
			// reject(err);
		});
	});
}

/* Send alerts to Syslog and patch alerts on Microsoft Graph */
async function sendAndPatchAlerts(securityAlerts){
	try{
		for (var i = 0; i <securityAlerts.value.length; i++){
		    syslogSend(securityAlerts.value[i]);
		    await patchAlert(securityAlerts.value[i]);
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
		if (securityAlerts.value.length != 0) {
			await sendAndPatchAlerts(securityAlerts);
		}
		else {
			moreAlerts = false;
			console.log("no more alerts!");
		}
	} while (moreAlerts);
}

function syslogSend(alert) {
	console.log("Sending Syslog " + alert.id);
	var syslog = require("syslog-client");

	// Options for syslog connection
	var options = {
		syslogHostname: GetEnvironmentVariable(SYSLOG_HOSTNAME) || "SyslogForwarder",
		transport: GetEnvironmentVariable(SYSLOG_TRANSPORT),
		port: parseInt(GetEnvironmentVariable(SYSLOG_PORT))
	};

	var client = syslog.createClient( GetEnvironmentVariable(SYSLOG_SERVER) , options);

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
