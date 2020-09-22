const ews = require('ews-javascript-api');

const EWS 	= process.env.EWS;
const USER 			= process.env.USER;
const PASSWORD 	= process.env.PWD;

var ewsAuth = require("ews-javascript-api-auth");
ews.ConfigurationApi.ConfigureXHR(new ewsAuth.ntlmAuthXhrApi(USER,PASSWORD));
ews.ConfigurationApi.SetXHROptions({rejectUnauthorized : true});

if( !!!process.env.EWS || !!!process.env.USER || !!!process.env.PWD){
	console.error('Provide EWS, USER and PWD tot the command');
	console.error('For example: EWS=https://ews.domain.com/EWS/Exchange.asmx USER=xxx@yyy.zzz PWD=xyz node ./index.js');
	return;
}

var service 					= new ews.ExchangeService(ews.ExchangeVersion.Exchange2016);
service.Credentials 	= new ews.WebCredentials(USER,PASSWORD);
service.Url 					= new ews.Uri(EWS);
service.TraceEnabled 	= false;

// DOCS on props of EmailMessage object
// http://ews-javascript-api.github.io/api/classes/core_serviceobjects_items_emailmessage.emailmessage.html
async function readEmail(notification){
	const EMAIL_META 				= notification.Events[0].itemId;

	// try to keep this mean and lean to speed things up
	var props = new ews.PropertySet(ews.BasePropertySet.IdOnly,[
		ews.EmailMessageSchema.Subject, 
		ews.EmailMessageSchema.TextBody, 
		ews.ItemSchema.Attachments, 
		ews.ItemSchema.HasAttachments,
	]);

  let email = await ews.Item.Bind(service,EMAIL_META,props);

  console.log('= = = = = = = = = = = = = = = = = = = = ');
  let subject = await email.Subject;
  console.log('LOADING DETAILS FOR EMAIL:', subject);
  console.log('= = = = = = = = = = = = = = = = = = = = ');
  
  // fetch all nested properties of email
  email.Load(props).then(() => {
  	console.log(email.TextBody.Text);
  	console.log('= = = = = = = = = = = = = = = = = = = = ');
    console.log("attachments:", email.HasAttachments);
  });
}

async function main() {
	const FOLDER_ID_INBOX = new ews.FolderId(ews.WellKnownFolderName.Inbox);
	const streamingSubscription = await service.SubscribeToStreamingNotifications([FOLDER_ID_INBOX], ews.EventType.NewMail);
	// Create a streaming connection to the service object, 
	// over which events are returned to the client.
	// Keep the streaming connection open for 30 minutes.
	let connection = new ews.StreamingSubscriptionConnection(service, 30);
	connection.AddSubscription(streamingSubscription);
	connection.OnNotificationEvent.push((o, a) => {
	  readEmail(a);
	});
	
	// max connection time is 30mins
	connection.OnDisconnect.push(() => {
	  // your logic for connecting again.
	  // you can use service SyncFolderItems 
	  // to sync changes after the app was offline
	  console.log('Disconnected, will attempt to reconnect');
	  connection.Open();
	});

	connection.Open();
}

main().then(() => {
		console.log("Subscribed to new incoming emails ...");
	}).catch( e => console.log( e ));

