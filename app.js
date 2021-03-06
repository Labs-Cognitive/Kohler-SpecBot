 //gkrishna1231@outlook.com
//Kri5hn@@
//33633
//https://specbot-b7a2.azurewebsites.net

var restify = require('restify');
var builder = require('botbuilder');
var LUIS = require('luis-sdk');
var moment = require('moment');
var unirest = require("unirest");
var path = require('path');
var request = require('request');
var fs = require('fs');

var json2xls = require('json2xls');

/*************  MODULE TWO  ********************/

var customer_name;
var pptname;
var pptid;

/*************  MODULE TWO  ********************/

var server = restify.createServer();

/*************  MODULE ONE  ********************/

server.listen(process.env.port || process.env.PORT || 4000, function () {
    console.log("--------------------------------------------------------");
    console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " |  KohlerBot is running with the address : " + server.url);
    console.log("--------------------------------------------------------");
});

server.get('/public/*', restify.plugins.serveStatic({ directory: __dirname }));

var connector = new builder.ChatConnector({
    appId:'f136fda9-d324-4ef4-8d5a-29603381268b', 
    appPassword:'pyopHCHE315^mdrGVX43#+$'
});

var bot = new builder.UniversalBot(connector, {
    storage: new builder.MemoryBotStorage()
});
var model = 'https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/5732cf18-d71a-49ac-9e14-491e48c0efac?subscription-key=bcd3f91561a14411807569e6623ef01a&timezoneOffset=-360&q=';	

var recognizer = new builder.LuisRecognizer(model);
var dialog = new builder.IntentDialog({
    recognizers: [recognizer]
});

function getCardsAttachments1(session) {
    return [
        new builder.HeroCard(session)
        .buttons([
            builder.CardAction.imBack(session, 'Commercial', "Commercial"),
            builder.CardAction.imBack(session, 'Hospitality', "Hospitality"),
            builder.CardAction.imBack(session, 'High Rise Multi family', "High Rise Multi family"),
            builder.CardAction.imBack(session, 'Low Rise Multi family', "Low Rise Multi family"),
            builder.CardAction.imBack(session, 'Single Family Home', "Single Family Home"),
            builder.CardAction.imBack(session, 'Showroom', "Showroom")
        ]),
    ];
}

function getCardsAttachments2(session) {
    return [
        new builder.HeroCard(session)
        .buttons([
            builder.CardAction.imBack(session, 'Issue', "Issue"),
            builder.CardAction.imBack(session, 'Question', "Question"),
            builder.CardAction.imBack(session, 'Feedback', "Feedback")
        ]),
    ];
}

function getCardsAttachments4Yes_No(session) {
    return [
        new builder.HeroCard(session)
        .buttons([
            builder.CardAction.imBack(session, 'Yes', "Yes"),
            builder.CardAction.imBack(session, 'No', "No"),
        ])
    ];
}


var n = 0;
server.post('/api/messages', connector.listen());
bot.dialog('/', dialog);

 bot.on('conversationUpdate', function(message) {

    //console.log('in Conversation', message)

    if (message.membersAdded) {

        message.membersAdded.forEach(function(identity) {

            // Bot is joining conversation
            // - For WebChat channel you'll get this on page load.
            if (identity.id == message.address.bot.id) {
                var reply = new builder.Message()
                    .address(message.address)
                    //.suggestResponses(logTo//console)
                    .text("Welcome to Spec Bot!");

                bot.send(reply);
            }

        })
    }
}) 

dialog.matches('greetings', [
    function(session, args) {
        session.sendTyping();

        console.log("--------------------------------------------------------");
        console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " | Greetings Intent Matched");
        console.log("--------------------------------------------------------");
        //console.log("RequestToken", session.message.user.RequestToken);
        //console.log('access_token', session.message.user.token.access_token);
        console.log('access_token', session.message.user.token);
        session.send("Hi, I am Spec Bot - how can I help you?");
    }
]);

dialog.matches('capabilities', [
    function(session, args) {
        session.sendTyping();
        ////console.log(session);
        console.log("--------------------------------------------------------");
        console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " | Capabilities Intent Matched");
        console.log("--------------------------------------------------------");
        session.send(" I can help you create Presentation or Assist with an Incident ");

    }
]);

 dialog.matches('question', [
    function(session, args) {
        session.sendTyping();
        console.log("--------------------------------------------------------");
        console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " | Question Intent Matched");
        console.log("--------------------------------------------------------");
        if (args.entities[0].type == "Presentation::create") {
            session.send("Yes, I can certainly do that. Please enter presentation name.");
            //session.beginDialog('/waterfall', session);
        } else if (args.entities[0].type == "Presentation::open") { 
			//Opening a Presentation
            session.beginDialog('/openppt', session);
        }
    }
]); 

dialog.matches('pptname', [
    function(session, args) {
        session.sendTyping();
        console.log(session);
		console.log(args);
		//console.log(args.entities[0].entity);
		pptname = args.entities[0].entity;
        console.log("--------------------------------------------------------");
        console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " | pptName Intent Matched");
        console.log("--------------------------------------------------------");
       
if(args.entities[0].type=="Entertainment.Title"){
	   console.log('before roomname waterfall')
	   session.beginDialog('/waterfall', session);
		//session.send("Enter")
}
else{
	session.send("Enter correct title");
}
    }
]);


//WaterFall2
bot.dialog('/waterfall2', [
    function(session) {
        console.log(session.message.value, 'country waterfall2')
        session.sendTyping();
        if (session.message.value.type == 'countries') {
            session.userData.Customer.CountryCode = session.message.value.name;
            delete session.message.value;
        }
        if (session.message && session.message.value) {
            processSubmitAction(session, session.message.value);
            return;
        }

        unirest.get('http://appsbotdev.azurewebsites.net/api/Common/GetStates/' + session.userData.Customer.CountryCode)
            .headers({
                'CSRFToken': session.message.user.RequestToken,
                'Authorization': 'Bearer ' + session.message.user.token.access_token
            })
            .end(function(w) {
                var states = [];
                for (i = 0; i < JSON.parse(w.raw_body).length; i++) {
                    states[i] = {
                        'title': JSON.parse(w.raw_body)[i].StateName,
                        'value': JSON.parse(w.raw_body)[i].StateCode
                    }
                }

                var card = {
                    'contentType': 'application/vnd.microsoft.card.adaptive',
                    'content': {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.0",
                        "body": [{
                                "type": "TextBlock",
                                "text": "Select a State",
                                "size": "large",
                                "weight": "bolder"
                            },

                            {
                                "type": "Input.ChoiceSet",
                                "id": "name",
                                "style": "compact",
                                "choices": states
                            }

                        ],
                        "actions": [{
                            "type": "Action.Submit",
                            "title": "Okay",
                            'data': {
                                'type': 'states'
                            }
                        }]
                    }
                };

                var msg = new builder.Message(session).addAttachment(card);
                session.send(msg);

            })
    }
])



//Excel Waterfall
bot.dialog('/excel', [
    function(session) {
		console.log("Excel Waterfalll")
        session.sendTyping();
        //session.userData.Customer.StateCode = session.message.value.name;

        builder.Prompts.attachment(session, "Please upload the excel sheet with **SKU**s, **Room** names and more details")

        var msg = new builder.Message(session)
            .attachments([{
                name: ' You can use this template file ',
                contentType: 'application/octet-stream',
                contentUrl: 'https://demodisks703.blob.core.windows.net/kohlerspecbot/SpecDeckRoomsSampleTemplate.xlsx'
            }]);

        session.send(msg)

    }
])

//For Change in Country and State
bot.dialog('/waterfall1', [
    function(session) {
        console.log('waterfall1 started');
        session.sendTyping();
		console.log(session.message.value)
		console.log(session.message)
        if (session.message && session.message.value) {
			console.log('during processSubmitAction');
            processSubmitAction(session, session.message.value);
            return;
        }

        unirest.get('http://appsbotdev.azurewebsites.net/api/Common/GetCountries')
            .headers({
                'CSRFToken': session.message.user.RequestToken,
                'Authorization': 'Bearer ' + session.message.user.token.access_token
            })
            .end(function(r) {
                var countries = [];
                for (i = 0; i < JSON.parse(r.raw_body).length; i++) {
                    countries[i] = {
                        'title': JSON.parse(r.raw_body)[i].CountryName,
                        'value': JSON.parse(r.raw_body)[i].CountryCode
                    }
                }

                var card = {
                    'contentType': 'application/vnd.microsoft.card.adaptive',
                    'content': {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.0",
                        "body": [{
                                "type": "TextBlock",
                                "text": "Select a Country",
                                "size": "large",
                                "weight": "bolder"
                            },

                            {
                                "type": "Input.ChoiceSet",
                                "id": "name",
                                "style": "compact",
								"value": session.message.user.token.groupid,
                                "choices": countries
                            }

                        ],
                        "actions": [{
                            "type": "Action.Submit",
                            "title": "Okay",
                            'data': {
                                'type': 'countries'
                            }
                        }]
                    }
                };
                //console.log('during attachment');
                var msg = new builder.Message(session).addAttachment(card);
                session.send(msg);
            })
    }
])

//For Change in Country and State


//Opening a Presentation
bot.dialog('/openppt', [
    function(session, args, results) {
        session.sendTyping();
        builder.Prompts.text(session, "Please enter Presentation name")
    },
    function(session, args, result) {
        session.sendTyping();
        builder.Prompts.text(session, "That Presentation name doesnot exit. Would you like to create one with this name?")
        var cards = getCardsAttachments4Yes_No();
        var reply = new builder.Message(session)
            .attachmentLayout(builder.AttachmentLayout.carousel)
            .attachments(cards);
        session.send(reply);
    },
    function(session, args, result) {
        //console.log("----------------------------------------------")
        //console.log(args)
        session.sendTyping();
        if (args.response == 'Yes') {
            session.beginDialog('/waterfall', session)
        } else {
            builder.Prompts.text(session, "Please enter presentation name");
        }
    }
])

// Creating a Presentation

bot.dialog('/waterfall', [


    function(session, args) {
        session.sendTyping();
        //console.log("-------------------------------------")
        //session.userData.Name = args.response;
        builder.Prompts.text(session, "Thank You. Select a Project Type from below.");

        var cards = getCardsAttachments1();
        var reply = new builder.Message(session)
            .attachmentLayout(builder.AttachmentLayout.carousel)
            .attachments(cards);
        session.send(reply);

    },
    function(session, args, results) {
		console.log('in select of Proect type')
        session.sendTyping();
        switch (args.response) {
            case 'Hospitality':
                session.userData.ProjectType = 'Proj1';
                break;
            case 'Commercial':
                session.userData.ProjectType = 'Proj2';
                break;
            case 'High Rise Multi family':
                session.userData.ProjectType = 'Proj3';
                break;
				case 'Low Rise Multi family':
                session.userData.ProjectType = 'Proj4';
                break;
            case 'Single Family Home':
                session.userData.ProjectType = 'Proj5';
                break;
            case 'Showroom':
                session.userData.ProjectType = 'Proj6';
                break;
            default:
                //builder.Prompts.text(session, "Thanks. Have a great day");
        }
        if (args.response == 'Commercial' || 'Hospitality' || 'High Rise Multi family' || 'Single Family Home' || 'Showroom') {
            session.beginDialog('/usergroup', session)
        } else {
            builder.Prompts.text(session, "Thanks. Have a great day");
        }

    }
])
bot.dialog('/usergroup', [

    function(session) {
        session.sendTyping();
        //console.log("#################################################################################################")
        //console.log(session.userData, 'userData');
        //console.log(session.message.user.RequestToken, 'userData');
        //console.log(session.message.user.token.access_token, 'userData');
		console.log(session.userData.ProjectType);
        if(session.message && session.message.value) {
            processSubmitAction(session, session.message.value);
            return;
        }

        //API for getting usergroup details
        unirest.get('http://appsbotdev.azurewebsites.net/api/GroupManagement/GetGroups')
            .headers({
                'CSRFToken': session.message.user.RequestToken,
                'Authorization': 'Bearer ' + session.message.user.token.access_token
            })
            .end(function(r) {
                if (r.ok) {
					var len = JSON.parse(r.raw_body).length;
                    var choices = [];
                    for (i = 0; i < len; i++) {
                        choices[i] = {
                            'title': JSON.parse(r.raw_body)[i].Name,
                            'value': JSON.parse(r.raw_body)[i].id
                        }
                    }
console.log(session.message.user.token.groupid, 'groupid');
console.log(choices, 'choices');
                    var card = {
                        'contentType': 'application/vnd.microsoft.card.adaptive',
                        'content': {
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "type": "AdaptiveCard",
                            "version": "1.0",
                            "body": [{
                                    "type": "TextBlock",
                                    "text": "Select User Group",
                                    "size": "large",
                                    "weight": "bolder"
                                },

                                {
                                    "type": "Input.ChoiceSet",
                                    "id": "name",
                                    "style": "compact",
									"value": session.message.user.token.groupid,
                                    "choices": choices
                                }

                            ],
                            "actions": [{
                                "type": "Action.Submit",
                                "title": "Okay",
                                'data': {
                                    'type': 'usergroup'
                                }
                            }]
                        }
                    };

                    var msg = new builder.Message(session).addAttachment(card);
                    session.send(msg);
                } else {
                    session.send("API's are not Authorized");
                }
            })

    }
]);



function processSubmitAction(session, value) {
    console.log(session.userData, 'userData at processSubmitAction', value.type);
    var defaultErrorMessage = 'Please complete all the search parameters';
    switch (value.type) {
        case 'usergroup':
            session.beginDialog('/custname', session)
            break;
        case 'countries':
            session.beginDialog('/waterfall2', session)
            break;
        case 'states':
            session.beginDialog('/excel', session)
            break;
        default:
            session.send(defaultErrorMessage);
    }
}

bot.dialog('/custname', [

    function(session, args, results) {
		console.log(session.userData.groupid)
		session.userData.groupid = session.message.value.name;
		console.log(session.message.value.name)
		console.log(session.userData.groupid)
        builder.Prompts.text(session, "Enter customer name");

    }, function(session, args, results) {

        builder.Prompts.text(session, "Enter the Room name");
		session.userData.customerName = session.message.text;
    }, function(session, args, results) {
		session.sendTyping();

	session.userData.RoomName = session.message.text;
	console.log('http://appsbotdev.azurewebsites.net/api/CustomerInfo/ByName/'+session.userData.customerName);
        //API to get Country and state details of customer based on Customer Name above
		unirest.get('http://appsbotdev.azurewebsites.net/api/CustomerInfo/ByName/'+session.userData.customerName)
            .headers({
                'CSRFToken': session.message.user.RequestToken,
                'Authorization': 'Bearer ' + session.message.user.token.access_token
            })
            .end(function(r) {
				console.log(r);
				console.log(r.raw_body);
				console.log(typeof(r.status));
				if(r.status != 200){
					session.send('Error: Enter the name Correctly');	
				}
				else{
                if (r.raw_body) {
                    var k = JSON.parse(r.raw_body)[0];
                    k.listpresentationCustomers = null;

                    session.userData.Customer = k;

                    builder.Prompts.confirm(session, "Country code **" + k.CountryCode +
                        "** and State code **" + k.StateCode + "** are being taken from your profile. Would you like to change them?");

                    var cards = getCardsAttachments4Yes_No();
                    var reply = new builder.Message(session)
                        .attachmentLayout(builder.AttachmentLayout.carousel)
                        .attachments(cards);
                    session.send(reply);
                } else {
                    session.send('I am unable to identify your details, try again provding the correct Customer Name.')
                }
				}
            })
    },
    function(session, args, results) {
        //console.log("==========================")
        //console.log(args.response, 'at SKUs')

        session.sendTyping();
        if (args.response == false) {

            builder.Prompts.attachment(session, "Please upload the excel sheet with **SKU**s, **Room** names and more details")

            var msg = new builder.Message(session)
                .attachments([{
                    name: ' You can use this template file ',
                    contentType: 'application/octet-stream',
                    contentUrl: 'https://demodisks703.blob.core.windows.net/kohlerspecbot/SpecDeckRoomsSampleTemplate.xlsx'
                }]);
            session.send(msg)

        } else {
            session.beginDialog('/waterfall1', session)
        }
    },
    function(session, args, results) {
        session.sendTyping();
		console.log(session.userData.Customer)
        // a REST API call for Creating Presentation
        if (path.extname(session.message.attachments[0].name) == '.xlsx') {
            unirest.post('http://appsbotdev.azurewebsites.net/api/PresentationSetup')
                .headers({
                    'CSRFToken': session.message.user.RequestToken,
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + session.message.user.token.access_token
                })
                .send({
                    "GroupId": session.message.user.token.groupid,
                    "CountryCode": session.userData.Customer.CountryCode,
                    "StateCode": session.userData.Customer.StateCode,
                    "Name": pptname,
                    "ProjectType": session.userData.ProjectType,
                    "customer": session.userData.Customer,//customer_name,
                    "CoverImagePath": null,
                    "ImageDetails": null,
                    "selectedImage": "https://stspecdeckdev.blob.core.windows.net/medialibrary/",
                    "BrandLogoPaths": null,
					"validForm":true
                })
                .end(function(response) {
                    console.log(response.raw_body);
                    console.log(response);
					if(response.ok){
                    unirest.post('http://appsbotdev.azurewebsites.net/api/GroupManagement/DisplayDefaultBrandCatalog')
                        .headers({
                            'CSRFToken': session.message.user.RequestToken,
                            'Content-Type': 'application/json',
                            'Authorization': 'Bearer ' + session.message.user.token.access_token
                        })
                        .send([session.message.user.token.groupid])
                        .end(function(t) {
                            BrandCatalogList = [];
                            t.raw_body.forEach(function(k) {
                                BrandCatalogList.push(k.BrandCode)
                            })
                            builder.Prompts.text(session, "Presentation **" + pptname + "** is created.")
							pptid=response.raw_body.id;
                            
							//API request to create Rooms based on the Attachement obtained above.
							request.post({
                                url: 'http://appsbotdev.azurewebsites.net/api/PresentationSetup/ImportProductsAndGetFailures',
                                headers: {
                                    'content-type': 'multipart/form-data',
                                    'CSRFToken': session.message.user.RequestToken,
                                    'Authorization': 'Bearer ' + session.message.user.token.access_token
                                },
                                formData: {
                                    importRoomProductObj: '{"RoomName":"'+session.userData.RoomName+'","presentationId": "' + response.raw_body.id + '","BrandCatalogsList":' + JSON.stringify(BrandCatalogList) + '}',
                                    importProductObj: 'null',
                                    excelFile: {
                                        value: request(session.message.attachments[0].contentUrl),
                                        options: {
                                            filename: session.message.attachments[0].name,
                                            contentType: '*/*'
                                        }
                                    }
                                }
                            }, function(error, responsee, body) {
                                if (error) throw new Error(error);
                                console.log(responsee)
                                console.log(JSON.parse(body).Message);
                                console.log(JSON.parse(body).SuccessRecordsCount);
                                console.log(JSON.parse(body).FailureRecordsCount);
                                if (body == JSON.stringify("Success")) {									
									   session.endDialog();
									   session.endConversation();
									   session.beginDialog('/Assistance1',session)
                                }
								else if(JSON.parse(body).Message == 'Export'){
									session.send('Processing of the import files is done.');
									session.send('**'+JSON.parse(body).SuccessRecordsCount + '** records are successfully imported.');
									session.send('**'+JSON.parse(body).FailureRecordsCount + '** records are failed to import.');
									   session.endDialog();
									   session.endConversation();
									   session.userData.success = JSON.parse(body).Jsonstring;
									   session.beginDialog('/partialimport',session)
								}
								else {
                                    session.send('Couldnot create Rooms.')
                                }
                            });
						})
				}
				else{
					session.send('There is a server error. Please try again.')	
				}

			}) 
		}
		else {
            
		session.beginDialog('/excellsheet', session)
			//session.send('Please upload an EXCEL Sheet');
			//session.beginDialog('/IncidentTitle', session);
			//session.endDialog();
			//session.endConversation();
		}
		}
]);		
		/* bot.dialog('/Assistance1',[
	
	function(session,args)
	{
		console.log("Inside Assistance")
		builder.Prompts.text(session, "Do you have any other assistance?")
		var cards = getCardsAttachments4Yes_No();
                                    var reply = new builder.Message(session)
                                        .attachmentLayout(builder.AttachmentLayout.carousel)
                                        .attachments(cards);
										session.send(reply);
		//session.send("Do you have any other assistance?")
	},
		function(session, args, results) {
        //console.log("----------------------------------------------------------", args.response)
        session.sendTyping();
        if (args.response == 'Yes') {
            session.send("I can help you create Presentation or Assist with an Incident")
			session.endDialog();
        session.endConversation(); 
        } else {

            builder.Prompts.text(session, "Thanks. Have a great day");
			session.endDialog();
        session.endConversation(); 
			
        }
		 

        
    }
		
]); */

//**************************************************************Excel Sheet After wrong upload***************************************

bot.dialog('/excellsheet',[
function(session, args, results) {
        //console.log("==========================")
        //console.log(args.response, 'at SKUs')

        session.sendTyping();
        

            builder.Prompts.attachment(session, "Problem reading **.xlsx file**! Please upload valid xlsx file")

            var msg = new builder.Message(session)
                .attachments([{
                    name: ' You can use this template file ',
                    contentType: 'application/octet-stream',
                    contentUrl: 'https://demodisks703.blob.core.windows.net/kohlerspecbot/SpecDeckRoomsSampleTemplate.xlsx'
                }]);
            session.send(msg)

        
    },
    function(session, args, results) {
        session.sendTyping();
		console.log(session.userData.Customer)
        // a REST API call for Creating Presentation
        if (path.extname(session.message.attachments[0].name) == '.xlsx') {
            unirest.post('http://appsbotdev.azurewebsites.net/api/PresentationSetup')
                .headers({
                    'CSRFToken': session.message.user.RequestToken,
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + session.message.user.token.access_token
                })
                .send({
                    "GroupId": session.message.user.token.groupid,
                    "CountryCode": session.userData.Customer.CountryCode,
                    "StateCode": session.userData.Customer.StateCode,
                    "Name": pptname,
                    "ProjectType": session.userData.ProjectType,
                    "customer": session.userData.Customer,//customer_name,
                    "CoverImagePath": null,
                    "ImageDetails": null,
                    "selectedImage": "https://stspecdeckdev.blob.core.windows.net/medialibrary/",
                    "BrandLogoPaths": null,
					"validForm":true
                })
                .end(function(response) {
                    console.log(response.raw_body);
                    console.log(response);
					if(response.ok){
                    unirest.post('http://appsbotdev.azurewebsites.net/api/GroupManagement/DisplayDefaultBrandCatalog')
                        .headers({
                            'CSRFToken': session.message.user.RequestToken,
                            'Content-Type': 'application/json',
                            'Authorization': 'Bearer ' + session.message.user.token.access_token
                        })
                        .send([session.message.user.token.groupid])
                        .end(function(t) {
                            BrandCatalogList = [];
                            t.raw_body.forEach(function(k) {
                                BrandCatalogList.push(k.BrandCode)
                            })
                            builder.Prompts.text(session, "Presentation **" + pptname + "** is created.")
							pptid=response.raw_body.id;

                            //API request to create Rooms based on the Attachement obtained above.
							request.post({
                                url: 'http://appsbotdev.azurewebsites.net/api/PresentationSetup/ImportProductsAndGetFailures',
                                headers: {
                                    'content-type': 'multipart/form-data',
                                    'CSRFToken': session.message.user.RequestToken,
                                    'Authorization': 'Bearer ' + session.message.user.token.access_token
                                },
                                formData: {
                                    importRoomProductObj: '{"RoomName":"'+session.userData.RoomName+'","presentationId": "' + response.raw_body.id + '","BrandCatalogsList":' + JSON.stringify(BrandCatalogList) + '}',
                                    importProductObj: 'null',
                                    excelFile: {
                                        value: request(session.message.attachments[0].contentUrl),
                                        options: {
                                            filename: session.message.attachments[0].name,
                                            contentType: '*/*'
                                        }
                                    }
                                }
                            }, function(error, response, body) {
                                if (error) throw new Error(error);
                                console.log(response)
                                console.log(body);
                                if (body == JSON.stringify("Success")) {
									//session.send("Sample Message");
									session.endDialog();
									session.endConversation();
									
                                   session.beginDialog('/Assistance1', session)
                                    //session.send('Do you need any other assistance?')
                                    
									console.log("*****************************inside room***********")
									}
									else if(JSON.parse(body).Message == 'Export'){
									session.send('Processing of the import files is done.');
									session.send('**'+JSON.parse(body).SuccessRecordsCount + '** records are successfully imported.');
									session.send('**'+JSON.parse(body).FailureRecordsCount + '** records are failed to import.');
									   session.endDialog();
									   session.endConversation();
									   session.userData.success = JSON.parse(body).Jsonstring;
									   session.beginDialog('/partialimport',session)
								}
									
							
                                 else {
                                    session.send('Couldnot create Rooms.')
                                }
                            });
						})
						
				}
				else{
					session.send('There is a server error. Please try again.')	
				}

			}) 
			
		}
		else {
            
		session.beginDialog('/excellsheet', session)
			//session.send('Please upload an EXCEL Sheet');
			//session.beginDialog('/IncidentTitle', session);
			/* session.endDialog();
			session.endConversation(); */
		}
		}
		]);
		
	bot.dialog('/partialimport',[
	
	function(session,args)
	{
	    session.sendTyping();
		session.send('Rooms are created with the succesfully imported details.');
		builder.Prompts.text(session, 'Do you want to download the failure imports?');
		var cards = getCardsAttachments4Yes_No();
        var reply = new builder.Message(session)
                        .attachmentLayout(builder.AttachmentLayout.carousel)
                        .attachments(cards);
		session.send(reply);
	},
	function(session,args)
	{
	    session.sendTyping();
		//call the API to download the failed import details.
		var json = [];

		for(i=0; i<JSON.parse(session.userData.success).length; i++){
			json[i] = {
				['Notes'] : ' ',
				["BRANDCATALOG"]: JSON.parse(session.userData.success)[i].BRANDCATALOG,
				["SKU"] : JSON.parse(session.userData.success)[i].SKU,
				["QUANTITY"] : JSON.parse(session.userData.success)[i].QUANTITY,
				["DESCRIPTION"] : JSON.parse(session.userData.success)[i].DESCRIPTION
			}
		}

		var xls = json2xls(json);

		fs.writeFileSync(__dirname+'/public/FailureRecords.xlsx', xls, 'binary');

			if(args.response == 'Yes'){
            var msg = new builder.Message(session)
                .attachments([{
                    name: 'Click here to get the details of failed imports ',
                    contentType: 'application/octet-stream',
                    contentUrl: 'http://kohlerspecdeck.azurewebsites.net/public/FailureRecords.xlsx'
                }]);
            session.send(msg)
			session.endDialog();
			session.endConversation();
			session.beginDialog('/Assistance1',session)
        } else {	
			session.endDialog();
			session.endConversation();
			session.beginDialog('/Assistance1',session)
		}
	}
	])

	bot.dialog('/Assistance1',[
	
	function(session,args)
	{
		console.log("Inside Assistance")
		builder.Prompts.text(session, "Do you want to export the Presentation?")
		var cards = getCardsAttachments4Yes_No();
                                    var reply = new builder.Message(session)
                                        .attachmentLayout(builder.AttachmentLayout.carousel)
                                        .attachments(cards);
										session.send(reply);
		//session.send("Do you have any other assistance?")
	},
		function(session, args, results) {
        //console.log("----------------------------------------------------------", args.response)
        session.sendTyping();
        if (args.response == 'Yes') {
            session.send("Please enter your mailid")
			session.endDialog();
        session.endConversation(); 
        } else {
					builder.Prompts.text(session, "Do you need any other assistance?")
		var cards = getCardsAttachments4Yes_No();
                                    var reply = new builder.Message(session)
                                        .attachmentLayout(builder.AttachmentLayout.carousel)
                                        .attachments(cards);
										session.send(reply);
           }
		 

        
    },
	function(session, args, results) {
        //console.log("----------------------------------------------------------", args.response)
        session.sendTyping();
        if (args.response == 'Yes') {
            session.send("I can help you create Presentation or Assist with an Incident")
			session.endDialog();
        session.endConversation(); 
		}
		else{
			builder.Prompts.text(session, "Thanks have a great day")
			session.endDialog()
			session.endConversation()
			
		}
	}
	
		
]);


//******************************************Excel sheet after wrong upload****************************************************************

//*********************************************INCIDENT*******************************************************************

dialog.matches('Incident', [
    function(session, args) {
        session.sendTyping();
        //console.log(args);
        console.log("--------------------------------------------------------");
        console.log(moment().format('MMMM Do YYYY, hhðŸ‡²ðŸ‡²ss a') + " | Incident Intent Matched");
        console.log("--------------------------------------------------------");
        if (args.entities[0].type == "IncidentCreate") { //API for creation
            //session.send("Yes, I can certainly do that. Please Enter Title.");
            session.beginDialog('/IncidentTitle', session);
        }

    }
]);
bot.dialog('/IncidentTitle', [

    function(session, args, results) {
        builder.Prompts.text(session, "Yes, I can certainly do that. Please Enter Title.");
    },
    function(session, args) {
        builder.Prompts.text(session, "Thank You. Select a Type from below.");

        var cards = getCardsAttachments2();
        var reply = new builder.Message(session)
            .attachmentLayout(builder.AttachmentLayout.carousel)
            .attachments(cards);
        session.send(reply);
    },
    function(session, args, results) {
        if (args.response == 'Issue' || 'Question' || 'Feedback') {
            session.beginDialog('/Incidentsubtype', session)
        } else {

            builder.Prompts.text(session, "Thanks. Have a great day");
        }

    }

]);
bot.dialog('/Incidentsubtype', [

    function(session) {
        if (session.message && session.message.value) {
            processSubmitAction1(session, session.message.value);
            return;
        }
        var card = {
            'contentType': 'application/vnd.microsoft.card.adaptive',
            'content': {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.0",
                "body": [{
                        "type": "TextBlock",
                        "text": "Select Sub Type",
                        "size": "large",
                        "weight": "bolder"
                    },

                    {
                        "type": "Input.ChoiceSet",
                        "id": "name",
                        "style": "compact",
                        "choices": [{
                                "title": "Catalogs",
                                "value": "Catalogs",
                                "isSelected": true
                            }, {
                                "title": "Copy center",
                                "value": "Copy center"
                            }, {
                                "title": "Cost summary",
                                "value": "Cost summary"
                            }, {
                                "title": "Customize presentation",
                                "value": "Customize presentation"
                            },
                            {
                                "title": "Email",
                                "value": "Email"
                            }, {
                                "title": "Export",
                                "value": "Export"
                            }, {
                                "title": "Media library",
                                "value": "Media library"
                            }, {
                                "title": "Notifications",
                                "value": "Notifications"
                            }, {
                                "title": "Presentations",
                                "value": "Presentations"
                            },
                            {
                                "title": "Rooms & Products",
                                "value": "Rooms & Products"
                            }, {
                                "title": "User administration",
                                "value": "User administration"
                            }, {
                                "title": "Others",
                                "value": "Others"
                            }

                        ]
                    }

                ],
                "actions": [{
                    "type": "Action.Submit",
                    "title": "Okay",
                    'data': {
                        'type': 'subtype'
                    }
                }]
            }
        };

        var msg = new builder.Message(session).addAttachment(card);
        session.send(msg);
    }
]);

function processSubmitAction1(session, value) {
    var defaultErrorMessage = 'Please complete all the search parameters';
    switch (value.type) {
        case 'subtype':
            //session.send(session.message.value.name)
            session.beginDialog('/IncidentDesc', session)
            break;

        default:
            session.send(defaultErrorMessage);
    }
};
bot.dialog('/IncidentDesc', [
    function(session, args, results) {
        builder.Prompts.text(session, "Please Enter your Description");
    },
    function(session, args, results) {
        builder.Prompts.text(session, "Thank you. Your Incident is Created");
        session.endDialog();
        session.endConversation();
    }

]);

/*************  MODULE ONE  ********************/


/*************  MODULE TWO  ********************/

/* dialog.matches('pptname', [
    function(session, args) {
        session.sendTyping();
        console.log(session);
		console.log(args);
		//console.log(args.entities[0].entity);
		pptname = args.entities[0].entity;
        console.log("--------------------------------------------------------");
        console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " | pptName Intent Matched");
        console.log("--------------------------------------------------------");
       
	if((args.entities[0].type=="Entertainment.Title") && (args.entities[1].type="Presentation:create"))
	{
	session.send("Only title");
	}
	   else
	{
	session.send("Enter correct title");
	}
	
	}
]); */

dialog.matches('projecttype', [
    function(session, args) {
        session.sendTyping();
        console.log(session);
        console.log("--------------------------------------------------------");
        console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " | prpjecttype Intent Matched");
        console.log("--------------------------------------------------------");
		 switch (args.response) {
            case 'Commercial':
                session.userData.ProjectType = 'Proj1';
                break;
            case 'Hospitality':
                session.userData.ProjectType = 'Proj2';
                break;
            case 'High Rise Multi family':
                session.userData.ProjectType = 'Proj3';
                break;
            case 'Single Family Home':
                session.userData.ProjectType = 'Proj4';
                break;
            case 'Showroom':
                session.userData.ProjectType = 'Proj5';
                break;
            default:
                //builder.Prompts.text(session, "Thanks. Have a great day");
        }
		if(args.entities[0]){
			console.log(args.entities[0].entity);
			 //args.entities[0].entity;
			    customer_name = args.entities[0].entity.charAt(0).toUpperCase()+ args.entities[0].entity.slice(1);
				console.log(customer_name)
			if ((args.response == 'Commercial' || 'Hospitality' || 'High Rise Multi family' || 'Single Family Home' || 'Showroom' ||'Low Rise Multi family') && (args.entities[0].type== "Entertainment.Person")) {
				console.log(args.entities[0].entity);
		switch (args.entities[1].entity) {
            case 'hospitality':
                session.userData.ProjectType = 'Proj1';
                break;
            case 'commercial':
                session.userData.ProjectType = 'Proj2';
                break;
            case 'high rise multi family':
                session.userData.ProjectType = 'Proj3';
                break;
				case 'low rise multi family':
                session.userData.ProjectType = 'Proj4';
                break;
            case 'single family home':
                session.userData.ProjectType = 'Proj5';
                break;
            case 'showroom':
                session.userData.ProjectType = 'Proj6';
                break;
            default:
                //builder.Prompts.text(session, "Thanks. Have a great day");
        }
            session.beginDialog('/countrycode', session)
        }
		  else{
			 session.send("Hit create customer api")
		 }  
			
		}
		else{
			if (args.response == 'Commercial' || 'Hospitality' || 'High Rise Multi family' || 'Single Family Home' || 'Showroom')  {
            session.beginDialog('/countrycode', session)
        }
		 
		  else if(args.response == 'Commercial' || 'Hospitality' || 'High Rise Multi family' || 'Single Family Home' || 'Showroom')
		 {
			 session.beginDialog('/countrycode', session)
		 }
		  else{
			 session.send("Hit create customer api")
		 } 
		}
        
 
    }
]);
bot.dialog('/countrycode',[
function(session,args)
{
	 console.log("*************************************************************************************************", customer_name);
	// console.log(JSON.stringify(args.entities))
        session.sendTyping();

        //API to get Country and state details of customer based on Customer Name above
        unirest.get('http://appsbotdev.azurewebsites.net/api/CustomerInfo/ByName/' + customer_name)
            .headers({
                'CSRFToken': session.message.user.RequestToken,
                'Authorization': 'Bearer ' + session.message.user.token.access_token
            })
            .end(function(r) {
				                    var k = JSON.parse(r.raw_body)[0];
									k.listpresentationCustomers = null;
                if (r.raw_body) {

                    session.userData.Customer = k;
                    builder.Prompts.confirm(session, "Country code **" + k.CountryCode +
                        "** and State code **" + k.StateCode + "** are being taken from your profile. Would you like to change them?");

                    var cards = getCardsAttachments4Yes_No();
                    var reply = new builder.Message(session)
                        .attachmentLayout(builder.AttachmentLayout.carousel)
                        .attachments(cards);
                    session.send(reply);
                } else {
                    session.send('I am unable to identify your details, try again provding the correct Customer Name.')
                }
            })
    },


function(session, args, results) {
        console.log("==========================")
        session.sendTyping();
        if (args.response == false) {
			session.beginDialog('/pptname', session)
			/* session.endDialog();
        session.endConversation(); */ 
		}
         else {
			 console.log('waterfall country assume');
            session.beginDialog('/country', session)
        }
		 
    }
]);

bot.dialog('/country', [
    function(session) {
        console.log('waterfall country started');
        session.sendTyping();
		console.log(session.message.value)
		console.log(session.message)
        if (session.message && session.message.value) {
			console.log('during processSubmitAction1');
            processSubmitAction1(session, session.message.value);
            return;
        }

        unirest.get('http://appsbotdev.azurewebsites.net/api/Common/GetCountries')
            .headers({
                'CSRFToken': session.message.user.RequestToken,
                'Authorization': 'Bearer ' + session.message.user.token.access_token
            })
            .end(function(r) {
                var countries = [];
                for (i = 0; i < JSON.parse(r.raw_body).length; i++) {
                    countries[i] = {
                        'title': JSON.parse(r.raw_body)[i].CountryName,
                        'value': JSON.parse(r.raw_body)[i].CountryCode
                    }
                }

                var card = {
                    'contentType': 'application/vnd.microsoft.card.adaptive',
                    'content': {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.0",
                        "body": [{
                                "type": "TextBlock",
                                "text": "Select a Country",
                                "size": "large",
                                "weight": "bolder"
                            },

                            {
                                "type": "Input.ChoiceSet",
                                "id": "name",
                                "style": "compact",
								"value": session.message.user.token.groupid,
                                "choices": countries
                            }

                        ],
                        "actions": [{
                            "type": "Action.Submit",
                            "title": "Okay",
                            'data': {
                                'type': 'countries1'
                            }
                        }]
                    }
                };
                //console.log('during attachment');
                var msg = new builder.Message(session).addAttachment(card);
                session.send(msg);
            })
    }
]);

bot.dialog('/state', [
    function(session) {
        console.log(session.message.value, 'country waterfall2')
        session.sendTyping();
        if (session.message.value.type == 'countries1') {
            session.userData.Customer.CountryCode = session.message.value.name;
            delete session.message.value;
        }
        if (session.message && session.message.value) {
            processSubmitAction1(session, session.message.value);
            return;
        }

        unirest.get('http://appsbotdev.azurewebsites.net/api/Common/GetStates/' + session.userData.Customer.CountryCode)
            .headers({
                'CSRFToken': session.message.user.RequestToken,
                'Authorization': 'Bearer ' + session.message.user.token.access_token
            })
            .end(function(w) {
                var states = [];
                for (i = 0; i < JSON.parse(w.raw_body).length; i++) {
                    states[i] = {
                        'title': JSON.parse(w.raw_body)[i].StateName,
                        'value': JSON.parse(w.raw_body)[i].StateCode
                    }
                }

                var card = {
                    'contentType': 'application/vnd.microsoft.card.adaptive',
                    'content': {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.0",
                        "body": [{
                                "type": "TextBlock",
                                "text": "Select a State",
                                "size": "large",
                                "weight": "bolder"
                            },

                            {
                                "type": "Input.ChoiceSet",
                                "id": "name",
                                "style": "compact",
                                "choices": states
                            }

                        ],
                        "actions": [{
                            "type": "Action.Submit",
                            "title": "Okay",
                            'data': {
                                'type': 'states1'
                            }
                        }]
                    }
                };

                var msg = new builder.Message(session).addAttachment(card);
                session.send(msg);

            })
    }
]);

function processSubmitAction1(session, value) {
    console.log(session.userData, 'userData at processSubmitAction', value.type);
    var defaultErrorMessage = 'Please complete all the search parameters';
    switch (value.type) {
        case 'usergroup':
            session.beginDialog('/custname', session)
            break;
        case 'countries1':
            session.beginDialog('/state', session)
            break;
        case 'states1':
            session.beginDialog('/pptname', session)
			/* session.endDialog();
        session.endConversation(); */
            break;
        default:
            session.send(defaultErrorMessage);
    }
}


	
	bot.dialog('/pptname',[
	function(session,args){
		builder.Prompts.text(session, 'Enter Presentation name');
		
		//session.send("Enter Room name");
		
	},
function(session,args){
	session.userData.Name = args.response;
	pptname=session.userData.Name;
		builder.Prompts.text(session, 'Enter room name');
		
		//session.send("Enter Room name");
		
	},
	function(session, args, results){
		console.log("/////////////////////////////////////////////////////////////////////////////////////////");
		console.log(args.response, 'at SKUs')
		session.userData.RoomName = session.message.text;
		
		 builder.Prompts.attachment(session, "Please upload the excel sheet with **SKU**s, **Room** names and more details")

            var msg = new builder.Message(session)
                .attachments([{
                    name: ' You can use this template file ',
                    contentType: 'application/octet-stream',
                    contentUrl: 'https://demodisks703.blob.core.windows.net/kohlerspecbot/SpecDeckRoomsSampleTemplate.xlsx'
                }]);
            session.send(msg)
	},
	
    function(session, args, results) {
        session.sendTyping();
		console.log(session.userData.Customer)
        // a REST API call for Creating Presentation
        if (path.extname(session.message.attachments[0].name) == '.xlsx') {
            unirest.post('http://appsbotdev.azurewebsites.net/api/PresentationSetup')
                .headers({
                    'CSRFToken': session.message.user.RequestToken,
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + session.message.user.token.access_token
                })
                .send({
                    "GroupId": session.message.user.token.groupid,
                    "CountryCode": session.userData.Customer.CountryCode,
                    "StateCode": session.userData.Customer.StateCode,
                    "Name": pptname,
                    "ProjectType": session.userData.ProjectType,
                    "customer": session.userData.Customer,//customer_name,
                    "CoverImagePath": null,
                    "ImageDetails": null,
                    "selectedImage": "https://stspecdeckdev.blob.core.windows.net/medialibrary/",
                    "BrandLogoPaths": null,
					"validForm":true
                })
                .end(function(response) {
        console.log(session.userData, 'Counter5 session', pptname, customer_name)
                    console.log(response);
					if(response.ok){
                    unirest.post('http://appsbotdev.azurewebsites.net/api/GroupManagement/DisplayDefaultBrandCatalog')
                        .headers({
                            'CSRFToken': session.message.user.RequestToken,
                            'Content-Type': 'application/json',
                            'Authorization': 'Bearer ' + session.message.user.token.access_token
                        })
                        .send([session.message.user.token.groupid])
                        .end(function(t) {
                            console.log(t.raw_body);
                            BrandCatalogList = [];
                            t.raw_body.forEach(function(k) {
                                console.log(k);
                                BrandCatalogList.push(k.BrandCode)
                            })
                            console.log(BrandCatalogList)

                            builder.Prompts.text(session, "Presentation **" + session.userData.Name + "** is created.")
							
							pptid = response.raw_body.id;
                            //API request to create Rooms based on the Attachement obtained above.
                            console.log('{"RoomName":"'+session.userData.RoomName+'","presentationId": "' + response.raw_body.id + '","BrandCatalogsList":"' + JSON.stringify(BrandCatalogList) + '"}');

                            request.post({
                                url: 'http://appsbotdev.azurewebsites.net/api/PresentationSetup/ImportProductsAndGetFailures',
                                headers: {
                                    'content-type': 'multipart/form-data',
                                    'CSRFToken': session.message.user.RequestToken,
                                    'Authorization': 'Bearer ' + session.message.user.token.access_token
                                },
                                formData: {
                                    importRoomProductObj: '{"RoomName":"'+session.userData.RoomName+'","presentationId": "' + response.raw_body.id + '","BrandCatalogsList":' + JSON.stringify(BrandCatalogList) + '}',
                                    importProductObj: 'null',
                                    excelFile: {
                                        value: request(session.message.attachments[0].contentUrl),
                                        options: {
                                            filename: 'session.message.attachments[0].name',
                                            contentType: '*/*'
                                        }
                                    }
                                }
                            }, function(error, response, body) {
                                if (error) throw new Error(error);
                                console.log(response)
                                console.log(body);
                                if (body == JSON.stringify("Success")) {
                                   session.endDialog();
								   session.endConversation();
                                    session.beginDialog('/Assistance2',session)
                                   
                                }
										else if(JSON.parse(body).Message == 'Export'){
									session.send('Processing of the import files is done.');
									session.send('**'+JSON.parse(body).SuccessRecordsCount + '** records are successfully imported.');
									session.send('**'+JSON.parse(body).FailureRecordsCount + '** records are failed to import.');
									   session.endDialog();
									   session.endConversation();
									   session.userData.success = JSON.parse(body).Jsonstring;
									   session.beginDialog('/partialimport',session)
								}								
								
								else {
                                    session.send('Couldnot create Rooms.')
                                }
                            });
                        })
				}
				else{
					session.send('There is a server error. Please try again.')	
				}

			}) 
		}
		else {
            session.beginDialog('/excellsheet1', session)
        }
		}
]);		
		
		/* bot.dialog('/Assistance3',[
	
	function(session,args)
	{
		console.log("Inside Assistance")
		builder.Prompts.text(session, "Do you have any other assistance?")
		var cards = getCardsAttachments4Yes_No();
                                    var reply = new builder.Message(session)
                                        .attachmentLayout(builder.AttachmentLayout.carousel)
                                        .attachments(cards);
										session.send(reply);
		//session.send("Do you have any other assistance?")
	},
		function(session, args, results) {
        //console.log("----------------------------------------------------------", args.response)
        session.sendTyping();
        if (args.response == 'Yes') {
            session.send("I can help you create Presentation or Assist with an Incident")
			session.endDialog();
        session.endConversation(); 
        } else {

            builder.Prompts.text(session, "Thanks. Have a great day");
			session.endDialog();
        session.endConversation(); 
			
        }
		 

        
    }
		
]); */

bot.dialog('/excellsheet1',[
function(session, args, results) {
        //console.log("==========================")
        //console.log(args.response, 'at SKUs')

        session.sendTyping();
        

            builder.Prompts.attachment(session, "Problem reading **.xlsx file**! Please upload valid xlsx file")

            var msg = new builder.Message(session)
                .attachments([{
                    name: ' You can use this template file ',
                    contentType: 'application/octet-stream',
                    contentUrl: 'https://demodisks703.blob.core.windows.net/kohlerspecbot/SpecDeckRoomsSampleTemplate.xlsx'
                }]);
            session.send(msg)

        
    },
    function(session, args, results) {
        session.sendTyping();
		console.log(session.userData.Customer)
        // a REST API call for Creating Presentation
        if (path.extname(session.message.attachments[0].name) == '.xlsx') {
            unirest.post('http://appsbotdev.azurewebsites.net/api/PresentationSetup')
                .headers({
                    'CSRFToken': session.message.user.RequestToken,
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + session.message.user.token.access_token
                })
                .send({
                    "GroupId": session.message.user.token.groupid,
                    "CountryCode": session.userData.Customer.CountryCode,
                    "StateCode": session.userData.Customer.StateCode,
                    "Name": pptname,
                    "ProjectType": session.userData.ProjectType,
                    "customer": session.userData.Customer,//customer_name,
                    "CoverImagePath": null,
                    "ImageDetails": null,
                    "selectedImage": "https://stspecdeckdev.blob.core.windows.net/medialibrary/",
                    "BrandLogoPaths": null,
					"validForm":true
                })
                .end(function(response) {
                    console.log(response.raw_body);
                    console.log(response);
					if(response.ok){
                    unirest.post('http://appsbotdev.azurewebsites.net/api/GroupManagement/DisplayDefaultBrandCatalog')
                        .headers({
                            'CSRFToken': session.message.user.RequestToken,
                            'Content-Type': 'application/json',
                            'Authorization': 'Bearer ' + session.message.user.token.access_token
                        })
                        .send([session.message.user.token.groupid])
                        .end(function(t) {
                            BrandCatalogList = [];
                            t.raw_body.forEach(function(k) {
                                BrandCatalogList.push(k.BrandCode)
                            })
                            builder.Prompts.text(session, "Presentation **" + pptname + "** is created.")
							pptid=response.raw_body.id ;
							

                            //API request to create Rooms based on the Attachment obtained above.
							request.post({
                                url: 'http://appsbotdev.azurewebsites.net/api/PresentationSetup/ImportProductsAndGetFailures',
                                headers: {
                                    'content-type': 'multipart/form-data',
                                    'CSRFToken': session.message.user.RequestToken,
                                    'Authorization': 'Bearer ' + session.message.user.token.access_token
                                },
                                formData: {
                                    importRoomProductObj: '{"RoomName":"'+session.userData.RoomName+'","presentationId": "' + response.raw_body.id + '","BrandCatalogsList":' + JSON.stringify(BrandCatalogList) + '}',
                                    importProductObj: 'null',
                                    excelFile: {
                                        value: request(session.message.attachments[0].contentUrl),
                                        options: {
                                            filename: session.message.attachments[0].name,
                                            contentType: '*/*'
                                        }
                                    }
                                }
                            }, function(error, response, body) {
                                if (error) throw new Error(error);
                                console.log(response)
                                console.log(body);
                                if (body == JSON.stringify("Success")) {
									//session.send("Sample Message");
									session.endDialog();
									session.endConversation();
									
                                   session.beginDialog('/Assistance2', session)
                                    //session.send('')
                                    
									console.log("*****************************inside room***********")
									
							
                                } 
								
										else if(JSON.parse(body).Message == 'Export'){
									session.send('Processing of the import files is done.');
									session.send('**'+JSON.parse(body).SuccessRecordsCount + '** records are successfully imported.');
									session.send('**'+JSON.parse(body).FailureRecordsCount + '** records are failed to import.');
									   session.endDialog();
									   session.endConversation();
									   session.userData.success = JSON.parse(body).Jsonstring;
									   session.beginDialog('/partialimport',session)
								}
								
								else {
                                    session.send('Couldnot create Rooms.')
                                }
                            });
						})
						
				}
				else{
					session.send('There is a server error. Please try again.')	
				}

			}) 
			
		}
		else {
            
		session.beginDialog('/excellsheet1', session)
			//session.send('Please upload an EXCEL Sheet');
			//session.beginDialog('/IncidentTitle', session);
			/* session.endDialog();
			session.endConversation(); */
		}
		}
		]);
		
	bot.dialog('/Assistance2',[
	
	function(session,args)
	{
		console.log("Inside Assistance")
		builder.Prompts.text(session, "Do you want to export the Presentation?")
		var cards = getCardsAttachments4Yes_No();
                                    var reply = new builder.Message(session)
                                        .attachmentLayout(builder.AttachmentLayout.carousel)
                                        .attachments(cards);
										session.send(reply);
		//session.send("Do you have any other assistance?")
	},
		function(session, args, results) {
        //console.log("----------------------------------------------------------", args.response)
        session.sendTyping();
        if (args.response == 'Yes') {
            session.send("Please enter your mailid")
			session.endDialog();
        session.endConversation(); 
        } else {
					builder.Prompts.text(session, "Do you need any other assistance?")
		var cards = getCardsAttachments4Yes_No();
                                    var reply = new builder.Message(session)
                                        .attachmentLayout(builder.AttachmentLayout.carousel)
                                        .attachments(cards);
										session.send(reply);
           }
		 

        
    },
	function(session, args, results) {
        //console.log("----------------------------------------------------------", args.response)
        session.sendTyping();
        if (args.response == 'Yes') {
            session.send("I can help you create Presentation or Assist with an Incident")
			session.endDialog();
        session.endConversation(); 
		}
		else{
			builder.Prompts.text(session, "Thanks have a great day")
			session.endDialog()
			session.endConversation()
			
		}
	}
	
		
]);

dialog.matches('mailid', [
    function(session, args) {
        session.sendTyping();
        ////console.log(session);
        console.log("--------------------------------------------------------");
        console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " | mail id Intent Matched");
        console.log("--------------------------------------------------------");
		//session.send("Presentation is mailed to you");

						var tg = {
						"ToEmailids":  args.entities[0].entity, //get from User
						"CcEmailids": "",
						"Subject": pptname,
						"PresentationName": pptname, //Name we got from user
						"PresentationId": pptid, //get PresentationID after Presentation is created.
						"Body": "Hello, \n\n  \n\n Here is the download link and password: \n\n  \n\n Please download and save the presentation onto your computer within 14 days of this email. After 14 days this link will expire.\n\n"
						}
		unirest.post('http://appsbotdev.azurewebsites.net/api/SharePresentation/Save')
		.headers({
			'CSRFToken': session.message.user.RequestToken,
			'Content-Type': 'application/json',
			'Authorization': 'Bearer ' + session.message.user.token.access_token
			})
		.send(tg)
		.end(function(result2){
			console.log(result2)
			console.log(result2.raw_body)
			session.send("Presentation **"+ pptname+"** is mailed to you. You'll receive a mail in a minute.")
			session.endDialog();
		session.endConversation();
		 session.beginDialog('/afterpptmail', session)
			})
		

    }
]);
bot.dialog('/afterpptmail',[


function(session,args)
	{
		console.log("Inside Assistance")
		builder.Prompts.text(session, "Do you need any other assistance?")
		var cards = getCardsAttachments4Yes_No();
                                    var reply = new builder.Message(session)
                                        .attachmentLayout(builder.AttachmentLayout.carousel)
                                        .attachments(cards);
										session.send(reply);
		
	},
		
	function(session, args, results) {
        //console.log("----------------------------------------------------------", args.response)
        session.sendTyping();
        if (args.response == 'Yes') {
            session.send("I can help you create Presentation or Assist with an Incident")
			session.endDialog();
        session.endConversation(); 
		}
		else{
			builder.Prompts.text(session, "Thanks have a great day")
			session.endDialog()
			session.endConversation()
			
		}
	}
	
		
]);

/*************  MODULE TWO  ********************/

dialog.matches('thankyou', [
    function(session, args) {
        session.sendTyping();
        //console.log(session, args);
        //console.log("--------------------------------------------------------");
        //console.log(moment().format('MMMM Do YYYY, hh:mm:ss a') + " | Thank You Intent Matched");
        //console.log("--------------------------------------------------------");
        session.send("Thank you. Have a great day!");
    }
]);

dialog.onDefault(builder.DialogAction.send("Sorry, I have trouble understanding you."))