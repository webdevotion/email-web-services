# Monitor an Inbox via EWS

Using [Exchange Web Service](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange) or EWS we can subscribe to notifications. This script subscribes to incoming events on Inbox and was developed against an Exchange 2016 server.

This script has not been used in production and is nothing more than a working example for you to tinker with.

## Instructions

Run this script from the command line but first edit the values of EWS, USER, and PWD to reflect your situation.

`EWS=https://ews.domain.com/EWS/Exchange.asmx USER=xxx@yyy.zzz PWD=xyz node ./index.js`

### Example Output

This email is a reply on an email that was sent. It's up to you to parse the body
into useful chunks. 

```
EWS=https://ews.domain.com/EWS/Exchange.asmx USER=xxx@yyy.zzz PWD=xyz node ./index.js

= = = = = = = = = = = = = = = = = = = =
LOADING DETAILS FOR EMAIL: Re: new email event stream
= = = = = = = = = = = = = = = = = = = =
text: Hi,

Is this a working example?




From: Jane Doe <jane.doe@myorg.tld>
Date: Tuesday, 22 September 2020 at 09:24
To: John Doe <john.doe@myorg.tld>
Subject: new email event stream

Hi

Please find an example in attachment.

attachments: true

disconnected, will attempt to reconnect
```
