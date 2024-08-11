# Outlook Email Bot

**Outlook Email Bot** is a tool for sending mass emails using Nodemailer and an Excel spreadsheet to manage recipients. With this bot, you can automate sending personalized emails to a list of contacts.

## Requirements

- Node.js (version 14 or higher)
- Node.js dependencies (listed in `package.json`)

## Installation

1. Install dependencies

```bash
npm install
```

2. Create the configuration file:

Rename config.example.json to config.json and fill in your information. An example of config.json is below:



```json
{
    "subject": "email title/subject test",
    "text": "test message ${mrs} ${name} send to email ${email}",
    "email": "testmail@gmail.com",
    "password": "pass",
    "docXSLX": "teste.xlsx",
    "host": "hotmail"
}
```

3. Prepare the Excel file:

1. Create an .xlsx file with the following headers: email, name, mrs.


2. Save the file in the same folder as index.js with the name specified in docXSLX in the configuration file.


# Usage

Run the script with the command:

```bash
npm start
```

## Description of Configuration Fields

subject: The subject of the email.

text: The body of the email, where ${mrs}, ${name} and ${email} will be replaced with the values ​​from the corresponding columns in the spreadsheet.

email: Your email used to send emails.

password: Your password for email authentication.

docXSLX: The name of the Excel file containing the list of emails and personal data.

host: The email provider (gmail, hotmail or yahoo).

# Messages in the Console

The bot will display messages in the console to indicate the status of operations:

Green: Email sent successfully.
Blue: Welcome message and instructions.
Red: Errors sending emails or problems with the configuration file.

Notes:

App Permissions: For services like Gmail, you may need to allow less secure apps or use an app password if two-step verification is turned on. Don't forget to disable two-step verification.
Security: Be sure to secure your configuration file, especially if you are sharing code or working in a public repository.

# License

This project is licensed under the MIT License.