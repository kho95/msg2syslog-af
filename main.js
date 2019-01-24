const MicrosoftGraph = require("@microsoft/microsoft-graph-client");

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

var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
        done(null, "PassInAccessTokenHere"); //first parameter takes an error if you can't get an access token
    }
});