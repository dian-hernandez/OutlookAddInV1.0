// Write to files
const fs = require('fs')

// Connect to email server and retrieve emails as a stream
const Imap = require('imap');

// For sending mail
const nodemailer = require('nodemailer');
//const xoauth2 = require('xoauth2');

// To parse data into readable format
const {simpleParser} = require('mailparser');
const { parse } = require('path/posix');

// For converting object
const { stringify } = require('querystring');

// Holds email data collected
var convData = null;

// Access to Office 365 server (obtaining email data)
const imapConfig = 
{
    user: 'sample@.com',
    password: 'password',
    host: 'outlook.office365.com' ,
    port: 993,
    tls: true
};

// Access to Office 365 server (sender)
let transporter = nodemailer.createTransport(
{
    host: 'smtp-mail.outlook.com',
    port: 587,
    auth:
    {
        user: 'sample@.com',
        pass: 'password'
    },
});

var mailOptions = 
{
    from: 'sample@.com',
    to: 'sample@.com',
    subject: 'Reported Email',
    attachments:
    {
        filename: 'data.txt',
        path: ''
    }
};

// Get Email contents
const getEmail = () => 
{
    try
    {
        const imap = new Imap(imapConfig);
    imap.once('ready', () => 
    {
      imap.openBox('INBOX', false, (err, results) => 
      { 
        imap.search(['UNSEEN'], (err, results) => 
        //imap.search(['UNSEEN', ['SINCE', new Date()]], (err, results) => 
        {
          const f = imap.fetch(results, {bodies: ''});
          f.on('message', msg => 
          {
            msg.on('body', stream => 
            {
              simpleParser(stream, async (err, parsed) => 
              {
                // All data such as {header, from, to, subject, message ID} = parsed
                // Parsed object is converted to string = convData
                convData = JSON.stringify(parsed);
                var tempFile = 'data.txt';

                fs.writeFile(tempFile, convData, err => 
                {
                    sendEmail();
                    if (err) 
                    {
                        console.error(err)
                        return
                    } 
                })
                 
              });
            });
          });
          f.once('error', ex => {
            return Promise.reject(ex);
          });
          f.once('end', () => 
          {
            console.log('Done fetching all messages!');
            imap.end();
          });
        });
      });
    });
        imap.once('error', err => 
        {
            console.log(err);
        });

        //Disconnect from server
        imap.once('end', () => {
        });

        imap.connect();

    }
    // Alert that error occurred while attempting to read
    catch (ex)
    {
        console.log('An error occurred');
    }
};

// Send  email out to address
function sendEmail ()
{
    // Obtains mailOptions specified 
   transporter.sendMail(mailOptions, function (err, info)
    {
        if (err)
        {
            console.log(err);
            return;
        }
    });
};

getEmail();