"use strict";
var builder = require("botbuilder");
var restify = require('restify');
var AuthenticationContext = require('adal-node').AuthenticationContext;
var Promise = require('es6-promise').Promise;
var _ = require('lodash');

//=========================================================
// ADAL Configuration
//=========================================================

var adalConfig = {
    'clientId' : '<your client id>', // The client Id retrieved from the Azure AD App
    'clientSecret' : '<your client secret>', // The client secret retrieved from the Azure AD App
    'authorityHostUrl' : 'https://login.microsoftonline.com/', // The host URL for the Microsoft authorization server
    'tenant' : '<your tenant id or name>>', // The tenant Id or domain name (e.g mydomain.onmicrosoft.com)
    'redirectUri' : process.env.REDIRECT_URI, // This URL will be used for the Azure AD Application to send the authorization code.
    'resource' : 'https://<your tenant name>.sharepoint.com', // The resource endpoint we want to give access to (in this case, SharePoint Online)
}

adalConfig.authorityUrl = adalConfig.authorityHostUrl + adalConfig.tenant;
adalConfig.templateAuthzUrl =  adalConfig.authorityUrl +
                        '/oauth2/authorize?response_type=code&client_id=' + // Optionally, we can get an Open Id Connect id_token to get more info on the user (some additional parameters are required if so https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-protocols-openid-connect-code)
                        adalConfig.clientId + 
                        '&state=<state>&resource=' + 
                        adalConfig.resource + 
                        '&response_mode=form_post' + //We want response as POST http request (see callback to see why)
                        '&redirect_uri=' + adalConfig.redirectUri  // If not specified, the adalConfigured reply URL of the Azure AD App will be used 
                    
//=========================================================
// Bot Setup
//=========================================================
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

var bot = new builder.UniversalBot(connector);

// Make sure you add code to validate these fields
var luisAppId = process.env.LUIS_APP_ID;
var luisAPIKey = process.env.LUIS_API_KEY;
var luisAPIHostName = process.env.LUIS_API_HOSTNAME || 'westus.api.cognitive.microsoft.com';

const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v1/application?id=' + luisAppId + '&subscription-key=' + luisAPIKey;

// Main dialog with LUIS
var recognizer = new builder.LuisRecognizer(LuisModelUrl);
var intents = new builder.IntentDialog({ recognizers: [recognizer] })

.matches('FindDocuments', () => {

    // TODO
})
.onDefault((session) => {
    session.send('Sorry, I did not understand \'%s\'.', session.message.text);
});


//=========================================================
// Server Setup (Restify)
//=========================================================
var port = process.env.port || process.env.PORT || 3978; // The port number is automatically assigned by Azure if hosted via the Web Application
var server = restify.createServer();
server.use(restify.bodyParser()); // To be able to get the authorization code (req.params.code)

server.listen(port, () => {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// This route is the endpoint for our bot (i.e which you put when you registrer your bot)
server.post('/api/messages', connector.listen()); 

// Create a route for the Azure AD App callback
// Be careful here: if you specfify a GET request for the OAuth callback, IISNODE will interpret the response as a static file due to the query string parameters instead of redirect it to the correct node js server route.
// To avoid modify the web.config, use a POST request instead
server.post('/api/oauthcallback', (req, res, next) => {

    // Get the authorization code from the Azure AD application
    var authorizationCode = req.params.code;
    if(authorizationCode) {

        acquireTokenWithAuthorizationCode(authorizationCode).then((response) => {

            // Add the state to the response to validate the CSRF scenario
            // The state has two utilities here:
            // - Reconnect with the bot to continue dialog
            // - Avoid CRSF attacks
            var state = req.params.state;
            if (state) {
                
                var address = JSON.parse(state);
                response.state = state;

                // Continue the dialog with the bot. Be careful, beginDialog" starts a new conversation.
                // We use the state parameter to save the address and be able to reconnect with the bot after authentication
                // Special thanks to this blog post https://dev-hope.blogspot.ca/2016/09/google-oauth-using-nodejs-and-microsoft.html
                // https://docs.botframework.com/en-us/node/builder/chat/UniversalBot/#navtitle ==> See paragraph "Saving Users Address"
                bot.beginDialog(address, "/oauth-success", response);
            }
        
            var body = '<html><body>Authentication succeeded! You can now close this tab</body></html>';
            res.send(200, body, { 'Content-Length': Buffer.byteLength(body), 'Content-Type': 'text/html' });
            res.end();

        }).catch((errorMessage) => {
            
            var body = '<html><body>' + errorMessage + '</body></html>';
            res.send(200, body, { 'Content-Length': Buffer.byteLength(body), 'Content-Type': 'text/html' });
            res.end();
        });
        
    } else {

        var body = '<html><body>Something went wrong, we didn\'t get an authorization code</body></html>';
        res.send(200, body, { 'Content-Length': Buffer.byteLength(body), 'Content-Type': 'text/html' });
        res.end();
    }
});

//=========================================================
// Bot authorization delegation middleware
//=========================================================
var getAuthorization = (session, args, next) => {

    // User is not already signed-in
    if (!session.privateConversationData['accessToken']) {

        // Set the arbitrary state as the current session address
        var stateToken = encodeURIComponent(JSON.stringify(session.message.address))
        var authorizationUrl = adalConfig.templateAuthzUrl.replace('<state>', stateToken);

        var actionLabel = 'You need to sign in to Office 365 before playing with this bot!';
        var buttonLabel = 'Sign-in';
        var signInCard = null;

        // The Sign-In card is not supported by Microsoft Teams for now (23/01/2017)
        // https://msdn.microsoft.com/en-us/microsoft-teams/bots#cards-and-buttons
        if (session.message.address.channelId === "msteams") {

             var link = builder.CardAction.openUrl(session, authorizationUrl,buttonLabel)
             signInCard = new builder.ThumbnailCard(session)
             .title("Authorization required!")
             .text(actionLabel)
             .buttons([link]);

        } else {

            // Send sign-in card
            signInCard =  new builder.SigninCard(session)
                .text(actionLabel)
                .button(buttonLabel, authorizationUrl);        
        }

        var msg = new builder.Message(session).attachments([signInCard]);
        session.send(msg);

    } else {

        // If the user is  already signed-in, we check if the access token is expired
        var expiresOn = session.privateConversationData['expiresOn'];
        var refreshToken = session.privateConversationData['refreshToken']

        if (new Date(expiresOn) >= Date.now() ) {
            
            acquireTokenWithRefreshToken(refreshToken).then((response) => {

                // Refresh the token infos
                session.privateConversationData['accessToken'] = response.accessToken;
                session.privateConversationData['expiresOn'] = response.expiresOn;
                session.privateConversationData['refreshToken'] = response.refreshToken;

                next();

            }).catch((errorMessage) => {
                console.log(errorMessage);
            });
        } else {
            next();
        }             
    }
}

//=========================================================
// Bot Dialogs
//=========================================================
/*bot.dialog('/', [
    getAuthorization,
    (session) => {

        var keywords = session.message.text

        // Check if a a message has been typed
        if (keywords) {

            // For debugging purpose, we add an arbitrary command to reset the bot state (we also could have implement a logout mechanism).
            // Initially the native /deleteprofile command was used but it is not available in the Bot Framework v3 anymore.
            if (keywords === "reset") {
                session.privateConversationData = {};

                // Get back to the main dialog route and prompt for a sign in
                session.beginDialog("/");

            } else {

                var accessToken = session.privateConversationData['accessToken'];

                // Now we have the token so we can make authenticated REST all to SharePoint or Graph API endpoints.        
                doSearch(keywords, accessToken).then((res) => {

                    if (res.error) {
                        session.send("Error: %s", res.error.message.value);

                    } else {

                        var cards = [];
                        var results = res.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;

                        if (results.length > 0) {

                            // Format search results wit ha Thumbnail card
                            _.each(results, function(value) {

                                var title = _.find(value.Cells.results, function(o) { return o.Key === "Title"; }).Value;
                                var link = builder.CardAction.openUrl(session, 
                                    _.find(value.Cells.results, function(o) { return o.Key === "Path"; }).Value,
                                    'View')
                                var fileType = _.find(value.Cells.results, function(o) { return o.Key === "FileType"; }).Value;
                                var hitHighlightedSummary = _.find(value.Cells.results, function(o) { return o.Key === "HitHighlightedSummary"; }).Value;
                                hitHighlightedSummary = hitHighlightedSummary.replace(/<c0>|<\/c0>/g,"").replace(/<ddd\/>/g,"");
                                var elt = new builder.ThumbnailCard(session).title(title).text(_.unescape(hitHighlightedSummary)).subtitle("Type: " + fileType).buttons([link]);

                                cards.push(elt);       
                            });

                            // create reply with Carousel AttachmentLayout
                            var reply = new builder.Message(session)
                                .attachmentLayout(builder.AttachmentLayout.carousel)
                                .attachments(cards);

                            console.log(reply);
                            session.send(reply);
                            
                        } else {
                            session.send("Sorry, we didn't find anything for '\%s\'", keywords);
                        }
                    }
                });
            }
        }
    }]);   
*/
bot.dialog('/oauth-success', function (session, response) {

    // Check the state value to avoid CSRF attacks http://www.twobotechnologies.com/blog/2014/02/importance-of-state-in-oauth2.html
    if(encodeURIComponent(JSON.stringify(session.message.address)) !== encodeURIComponent(response.state)) {
        session.send("CSRF scenario detected. Closing the current conversation...");
        session.endDialog();
    } else {

        // Save the token for the current user and for this conversation only (privateConversationData)
        if (!session.privateConversationData['accessToken']) {
            
            session.privateConversationData['accessToken'] = response.accessToken;
            session.privateConversationData['expiresOn'] = response.expiresOn;
            session.privateConversationData['refreshToken'] = response.refreshToken;
        }

        session.send('Hi %s. What can I do for you today?', response.userName);

        // Get back to the main dialog route
        session.beginDialog("/");
    }
});

bot.dialog('/', getAuthorization, intents);

//=========================================================
// SharePoint utilities
//=========================================================
var doSearch = (query, accessToken) => {

    var p = new Promise((resolve, reject) => {

        var endpointUrl = adalConfig.resource + "/_api/search/query?querytext='" + query + "'"; 

        // Node fetch is the server version of whatwg-fetch
        var fetch = require('node-fetch');

        fetch(endpointUrl, {
            method: 'GET',
            headers: {
                // The APIs require an OAuth access token in the Authorization header, formatted like this: 'Authorization: Bearer <token>'. 
                "Authorization" :  "Bearer " + accessToken,
                // Needed to get the results as JSON instead of Atom XML (default behavior)
                "Accept" : "application/json;odata=verbose"
            }           
        }).then(function(res) {
            return res.json();
        }).then(function(json) {
            resolve(json);
        }).catch(function(err) {
            reject(err);
        });
    });

    return p;
}

//=========================================================
// ADAL Helper Methods
//=========================================================
var acquireTokenWithAuthorizationCode = (authorizationCode) => {

    var authenticationContext = new AuthenticationContext(adalConfig.authorityUrl);

    var p = new Promise((resolve, reject) => {

        authenticationContext.acquireTokenWithAuthorizationCode(
            authorizationCode,
            adalConfig.redirectUri, // This URL must be the same as the redirect_uri of the original request or the reply url of the Azure AD App. Otherwise, it will throw an error.
            adalConfig.resource,
            adalConfig.clientId, 
            adalConfig.clientSecret,
            (err, response) => {

                if (err) {
                    reject(errorMessage = 'error: ' + err.message + '\n');

                } else {
                    resolve({ 
                        userName: (response.givenName + " " + response.familyName),
                        accessToken: response.accessToken,
                        expiresOn: response.expiresOn,
                        refreshToken: response.refreshToken,
                    }); 
                }
            });
    });

    return p;
}

var acquireTokenWithRefreshToken = (refreshToken) => {

    var authenticationContext = new AuthenticationContext(adalConfig.authorityUrl);

    var p = new Promise((resolve, reject) => {

        authenticationContext.acquireTokenWithRefreshToken(
            refreshToken,
            adalConfig.clientId,
            adalConfig.clientSecret,
            adalConfig.resource,
            (err, response) => {

                if (err) {
                    reject(errorMessage = 'error: ' + err.message + '\n');

                } else {
                    resolve({ 
                        userName: (response.givenName + " " + response.familyName),
                        accessToken: response.accessToken,
                        expiresOn: response.expiresOn,
                        refreshToken: response.refreshToken,
                    }); 
                }
            });
    });

    return p;
}