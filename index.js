const nodemailer = require('nodemailer');
const maildetails = require("./config.json")
const XLSX = require('xlsx');
const fs = require('fs');
var workbook;

var emails = [];
var names = [];
var mrs = [];

async function main(){
    console.log('\x1b[32m%s\x1b[0m',
       `

       __________ ___________________    _____      _____    _________ _________ ___________   _____      _____  .___.____     
       \______   \\_____  \__    ___/   /     \    /  _  \  /   _____//   _____/ \_   _____/  /     \    /  _  \ |   |    |    
        |    |  _/ /   |   \|    |     /  \ /  \  /  /_\  \ \_____  \ \_____  \   |    __)_  /  \ /  \  /  /_\  \|   |    |    
        |    |   \/    |    \    |    /    Y    \/    |    \/        \/        \  |        \/    Y    \/    |    \   |    |___ 
        |______  /\_______  /____|    \____|__  /\____|__  /_______  /_______  / /_______  /\____|__  /\____|__  /___|_______ \
               \/         \/                  \/         \/        \/        \/          \/         \/         \/            \/       
        `)
    console.log('\x1b[34m%s\x1b[0m',`
        Welcome to Outlook Email Bot! 
        Send massive emails using this bot. 🎉
        Check readme.md for more details.
        `
    );
    
        
    var desc = maildetails.text
    var email = maildetails.email
    var password =maildetails.password 
    var subject = maildetails.subject
    var docXSLX = maildetails.docXSLX
    var host = maildetails.host
    if(desc == "" || desc == null || desc == undefined){
        return console.error('Fill in the "desc" field in the configuration file with the text that needs to be sent to continue.')
    }
    if(email == "" || email == null || email == undefined){
        return console.error('Fill in the "email" field in the configuration file with the email that will be used to send')
    }
    if(password == "" || password == null || password == undefined){
        return console.error('Fill in the "password" field in the configuration file with the password of the email that will be used to send')
    }
    if(subject == "" || subject == null || subject == undefined){
        return console.error('Fill in the "subject" field in the configuration file with the subject of the text that will be sent')
    }
    if(docXSLX == "" || docXSLX == null || docXSLX == undefined) {
        return console.error('Fill in the "docXSLX" field in the configuration file with the name of the .xlsx file that must be in the same folder as this project. Example: "emails.xlsx"')
    }
    if(host =="" || host == null || host == undefined){
        return console.error('Fill in "host" in the configuration file to define the type of service that will be used. For Google put it as "gmail", for Outlook put it as "hotmail", for Yahoo Mail put it as "yahoo".')
    }
    getWorderBook();
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    // Converter a aba em um JSON
    const data = XLSX.utils.sheet_to_json(sheet);
    
    // Iterar sobre os dados e preencher os arrays
    data.forEach(row => {
        if (row.email) emails.push(row.email);
        if (row.name) names.push(row.name);
        if (row.mrs) mrs.push(row.mrs);
    });

    for(let i = 0;i<emails.length;i++){
        await sendEmail(emails[i],subject,desc,mrs[i],names[i])
    }
}


function getWorderBook(){
    try{
        workbook = XLSX.readFile(maildetails.docXSLX);
    }catch(e){
        console.error('\x1b[31m%s\x1b[0m',maildetails.docXSLX+" no such file in this directory! Check if "+maildetails.docXSLX+" exists.")
        return false;
    }
}

async function sendEmail(to,subject,text,mrs,name){
// Configuração do transporte para enviar o email usando o Outlook
let transporter = nodemailer.createTransport({
    service: mai,
    auth: {
        user: maildetails.email, // Seu email do Outlook
        pass: maildetails.password // Sua senha do Outlook
    }
});
    let mailOptions = {
        from: maildetails.email, // Seu email
        to: to, // Email do destinatário
        subject: subject, // Assunto do email
        text: ((text.replaceAll("${mrs}", mrs))
        .replaceAll("${name}", name))
        .replaceAll("${email}", to) // Corpo do email
    };
    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            return console.log('\x1b[31m%s\x1b[0m',
                `Error to send email to ${to}: ${to === maildetails.email ? "Don't email yourself!" : error}`, 
            );
        } else {
            return console.log(
                '\x1b[32m%s\x1b[0m',
                `Email sent successfully to ${to}!`, 
            );
        }
    });
     
}

main();