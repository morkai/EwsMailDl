# EwsMailDl

Microsoft Exchange E-mail Sender and Attachment Downloader.

Tested with Exchange server version 2010 SP2.

## Attachment Downloader

1. Connects to the specified Exchange server.
2. Subscribes to streaming notifications of type `NewMail` in the specified folder.
3. Downloads attachments from all existing e-mails (e-mails with attachments
   and subjects matching the specified filters).
4. Downloads attachments from any new, matching e-mails.
5. If the subscription closes, repeat from step 3.
6. If any error occurs, crash (Windows Service Recovery should take care
   of restarting the service).

After all attachments are downloaded, the e-mail is deleted.

## E-mail Sender

1. Monitors the specified input path for files matching the `*.email` pattern.
2. Reads, parses and sends e-mail based on the file contents.

After the e-mail is sent, the file is deleted.

The `.email` file should have the following format:

```
<header-name>: <header-value>
Body:
<body>
```

Available headers are:

  * `Subject` - sets the `EmailMessage.Subject` property.

  * `Importance` - sets the `EmailMessage.Importance` property (`low`, `normal` or `high`).

  * `To`, `ToRecipients` - adds recipients to the `EmailMessage.ToRecipients` list.
    Multiple recipients should be separated by a comma.

  * `Cc`, `CcRecipients` - adds recipients to the `EmailMessage.CcRecipients` list.
    Multiple recipients should be separated by a comma.

  * `Bcc`, `BccRecipients` - adds recipients to the `EmailMessage.BccRecipients` list.
    Multiple recipients should be separated by a comma.

  * `From` - sets the `EmailMessage.From` property.
  
  * `ReplyTo` - sets the `EmailMessage.ReplyTo` property.

  * `Html` - `1` or `true` if the e-mail body is HTML. If not specified, the body will still
    be sent as HTML but all new lines `\n` will be replaced with `<br>` and spaces ` ` at
	the beginning of each line will be replaced with `&nbsp;`.

## Requirements

### .NET Framework

  * __Version__: 4.x
  * __Website__: http://www.microsoft.com/net
  * __Download__: http://www.microsoft.com/net/downloads

### EWS Managed API

  * __Version__: 2.x
  * __Website__: http://msdn.microsoft.com/en-us/library/dd633709(v=exchg.80).aspx
  * __Download__: http://www.microsoft.com/en-us/download/details.aspx?id=35371

## Usage

EwsMailDl can be run as a console application or a service application.

### Console

To run EwsMailDl as a console application, execute the following command:

```
EwsMailDl.exe <arguments>
```

where `<arguments>` is a list of the [configuration arguments](#configuration).

This mode is intended for testing purposes only.

### Service

To install EwsMailDl as a service, execute the following command:

```
EwsMailDl.exe /i <arguments>
```

where `<arguments>` is a list of the [configuration arguments](#configuration).

The serivce can be then started using the standard `net start` or `sc start`
commands.

This mode is intended for use in production. The created `EwsMailDl` service
should be configured to restart on failure (*Recovery* tab in the service's
properties), because it will crash on any error.

To uninstall the service, execute the following command:

```
EwsMailDl.exe /u
```

### Configuration

Configuration arguments are specified in the following format:

```
/<arg-name-1>="<arg-value-1>" /<arg-name-2>="<arg-value-2>" ...
```

for example:

```
/quas="q" /wex="w" /exort="e"
```

Available configuration arguments are:

  * `version` - a version of the Exchange server we are connecting to.
    Valid values are: `Exchange2010_SP1`, `Exchange2010_SP2` or `Exchange2013`.
    Defaults to `Exchange2010_SP2`.
  
  * `url` - an URL to the server's EWS. For example, if the server we're trying
    to connect to is `mail.example.com`, then the EWS URL should be:
    `https://mail.example.com/EWS/Exchange.asmx`.
  
  * `username` - a username of the e-mail account we're trying to connect to.
  
  * `password` - a password of the e-mail account we're trying to connect to.
  
  * `lifetime` - a number of minutes (between 1 and 30) the subscription
    notification is active on the server. Defaults to 30 minutes.
  
  * `folderName` - a name of the folder in the user's account we're going to
    be monitoring for e-mails. Can be a `WellKnownFolderName` or any other
    user-created folder. Defaults to `Inbox`.
  
  * `folderId` - an ID of the folder in the user's account we're going to be
    monitoring for e-mails. Optional. If specified, the `folderName` is not used.
  
  * `inputPath` - a path to a folder that should be monitored for `.email` files.
  
  * `savePath` - a path to a folder where the attachments should be downloaded to.
    Defaults to the current directory.
  
  * `subject` - a filter for the e-mails to download. Only e-mails with a subject
    containing the specified string will be taken into consideration (they must
    have attachments too).
    Can be specified multiple times.
    Multiple filters are concatenated using `OR`.
  
  * `timestamp` - determines whether to prepend `<unix-timestamp>@` to
    the downloaded attachment file names, where `<unix-timestamp>` is the e-mail's
    date received as a UNIX timestamp. For example, attachment named `Test.html`
    that arrived at 2014-01-02 12:00:00 GMT will be saved as `1388664000000@Test.html`.

#### Example

Running from console:

```
EwsMailDl.exe ^
  /url="http://mail.example.com/EWS/Exchange.asmx" ^
  /username="code1\foobar" ^
  /password="top $$$ecret" ^
  /folderName="Baz" ^
  /inputPath="C:/emails" ^
  /savePath="C:/attachments" ^
  /subject="FOO" ^
  /subject="BAR" ^
  /timestamp="1"
```

Sending a text e-mail:

```
Subject: Test text e-mail
To: someone@the.net
Body:
Hello World!
```

Sending an HTML e-mail:

```
Subject: Test HTML e-mail
To: someone@the.net
Html: 1
Body:
<h1>Hello World!</h1>
```

## License

This project is released under the
[MIT License](https://raw.github.com/morkai/EwsMailDl/master/license.md).
