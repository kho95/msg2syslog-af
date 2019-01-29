# Microsoft Graph Security REST API to Syslog

It's quite simple, this is a function which queries (via polling) security alerts from Microsoft Graph and forwards the alerts to a Syslog Server.

## To do
* Integration Tests
* Token fail fallback function
* Modify to run as Azure Functions App

## Requirements

* NodeJS and NPM Package Manager
* Access to Microsoft Graph
* A computer?

## Obtaining Access to Microsoft Graph API

Follow the Microsoft Guide for 'Get Access without a user' [available here](https://docs.microsoft.com/en-us/graph/auth-overview).

For this application to work, you will need:
* Tenant ID (Format: `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`)
* Client/Application ID (Format: `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`)
* Client Secret/Password (Format: `XXXXXXXXXXXXXXXXXXXXXXX`)

## Example `.env` file
```
tenant = XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
client_id = XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
scope = https://graph.microsoft.com/.default
client_secret = XXXXXXXXXXXXXXXXXXXXXXX
grant_type = client_credentials
time_interval = 60000
token_domain = [your-tenant-name].onmicrosoft.com
hostname = [SyslogForwarder]
ip_address = XXX.XXX.XXX.XXX
port = 514
transport = UDP
```

Note: `time_interval` controls the Graph API poll interval (in milliseconds).

## How to Run

1. Clone this repo `https://github.com/fong/msg2syslog.git` for HTTPS or `git@github.com:fong/msg2syslog.git` for SSH
2. Open command prompt and run `npm install`
3. Create a `.env` file on the root of the application directory (same level as `main.js`)
4. Load your Microsoft Graph Authorisation IDs and Keys into the `.env` file
5. Run the application with `npm start`

## Dependencies

[msgraph-sdk-javascript](https://github.com/microsoftgraph/msgraph-sdk-javascript)

[syslog-client](https://github.com/paulgrove/node-syslog-client)

