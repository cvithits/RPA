// See https://github.com/dialogflow/dialogflow-fulfillment-nodejs
// for Dialogflow fulfillment library docs, samples, and to report issues
'use strict';
 
// Import the service function and various response classes
const {
  dialogflow,
  actionssdk,
  Image,
  Table,
  Carousel,
} = require('actions-on-google');

const functions = require('firebase-functions');
const {WebhookClient} = require('dialogflow-fulfillment');
const {Card, Suggestion} = require('dialogflow-fulfillment');
const officegen = require('officegen')
const fs = require('fs')
 
process.env.DEBUG = 'dialogflow:debug'; // enables lib debugging statements
 
exports.dialogflowFirebaseFulfillment = functions.https.onRequest((request, response) => {
  const agent = new WebhookClient({ request, response });
  console.log('Dialogflow Request headers: ' + JSON.stringify(request.headers));
  console.log('Dialogflow Request body: ' + JSON.stringify(request.body));
 
  function welcome(agent) {
    agent.add(`Welcome to my agent!`);
  }
 
  function fallback(agent) {
    agent.add(`I didn't understand`);
    agent.add(`I'm sorry, can you try again?`);
  }

  function createDoc(info){
  // Create an empty Word object:
	let docx = officegen('docx');
 
// Officegen calling this function after finishing to generate the docx document:
docx.on('finalize', function(written) {
  console.log(
    'Finish to create a Microsoft Word document.'
  );
});
 
// Officegen calling this function to report errors:
docx.on('error', function(err) {
  console.log(err);
});

// Create a new paragraph:
let pObj = docx.createP();
 
pObj.addText(info);

let out = fs.createWriteStream('/tmp/response.docx');
 
out.on('error', function(err) {
  console.log(err);
});
 
// Async call to generate the output file:
docx.generate(out);
  }
  
  function sendMail(subject, info, email){
     	createDoc(info);
        const nodemailer = require('nodemailer');
        const transporter = nodemailer.createTransport({
          host: "smtp.office365.com",
          port: 587,
          secure: false, // upgrade later with STARTTLS
          auth: {
            user: "khurram.khan@in2ittech.com",
            pass: "Admin@123"
          }
        });
    
        var mailOptions = {
          from: 'khurram.khan@in2ittech.com',
          to: email, //receiver email 
          subject: subject,
          text: info,
            attachments: [
          {   // utf-8 string as an attachment
              path: '/tmp/response.docx'
          }
              ]
        };
    
        transporter.sendMail(mailOptions, function (error, info) {
          if (error) {
            console.log('Error with:'+error);
          } else {
            console.log('Email sent: ' + info.response);
          }
        });
  }
  
  function sendMailHandler(agent){
      sendMail("REPLY TO GET_WATER_LEVEL",agent.parameters.info, agent.parameters.email);
      agent.add('Replied to subject: '+"GET_WATER_LEVEL"+" To "+agent.parameters.email+" "+agent.parameters.info);
  }
  
/*function getWaterLevelHandler(agent){
    agent.add('Forward the request to RPA request: '+"GET_WATER_LEVEL"+' parameter: '+ agent.parameters.address+' from email: '+agent.parameters.email);
  	sendMail("GET_WATER_LEVEL",agent.parameters.address, agent.parameters.email);
  }*/
 
  let intentMap = new Map();
  intentMap.set('Default Welcome Intent', welcome);
  intentMap.set('Default Fallback Intent', fallback);
  intentMap.set('send email Intent', sendMailHandler);
  //intentMap.set('get water level intent', getWaterLevelHandler);
  // intentMap.set('your intent name here', yourFunctionHandler);
  // intentMap.set('your intent name here', googleAssistantHandler);
  agent.handleRequest(intentMap);
});
