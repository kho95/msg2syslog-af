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

async function patchAlert(alert) {

	return new Promise(function (resolve, reject) {
		console.log("in patch alert");
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
			// reject(err);
		});
	});
}


async function sendAndPatchAlerts(securityAlerts){
	try{
		for (var i = 0; i <securityAlerts.value.length; i++){
			//console.log(securityAlerts.value[i]);
		    syslogSend(securityAlerts.value[i]);
		    await patchAlert(securityAlerts.value[i]);
		}
	}catch(err){
		console.log("Exception occured when accessing or sending alert data: " + err);
		console.log("Is securityAlerts object still what you expected?");
	}
	
}

async function getAlerts() {

	let moreAlerts = true;
	let _top = 1; //number of alerts to pull at a time
	let _skip = 0; //number of alerts to skip (offset) default to 0

	do{
		let securityAlerts = await getAlertsAPI(_top,_skip);
	
		console.log("in getalerts");

		if(securityAlerts.value.length != 0){
			for (var i=0; i<securityAlerts.value.length; i++){
				console.log(securityAlerts.value[i].id);
			}
			await sendAndPatchAlerts(securityAlerts);
		}
		else{
			moreAlerts = false;
			console.log("no more alerts!");
		}
	
	}while(moreAlerts);
}

function syslogSend(alertMessage) {
	console.log("in syslog send");
	// Initialising syslog
	var syslog = require("syslog-client");

	// Getting environment variables
	var SYSLOG_SERVER = "23.101.230.231";
	var SYSLOG_PROTOCOL = SYSLOG_PROTOCOL;
	var SYSLOG_HOSTNAME = SYSLOG_HOSTNAME;
	var SYSLOG_PORT = SYSLOG_PORT;

	// Options for syslog connection
	var options = {
		syslogHostname: "SyslogForwarder",
		transport: "UDP",
		port: 514
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
