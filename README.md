# Microsoft Graph Security REST API to Syslog

It's quite simple, this is a function which queries (via polling) alerts from MS Graph and forwards them to a Syslog Server.

## To do
* Connect up with syslog-client
* Move alert feed time interval to env file
* Rerequest token on expiration <- needs testing
* Modify to run as Azure Functions App

## Requirements

* NodeJS and NPM Package Manager
* Access to Microsoft Graph
* A computer?

## How to run

1. Clone this repo `https://github.com/fong/msg2syslog.git` for HTTPS or `git@github.com:fong/msg2syslog.git` for SSH
2. Open command prompt and run `npm install`
3. Run the application with `npm start`

## Dependencies

[msgraph-sdk-javascript](https://github.com/microsoftgraph/msgraph-sdk-javascript)

[syslog-client](https://github.com/paulgrove/node-syslog-client)

