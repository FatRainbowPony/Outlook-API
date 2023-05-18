# Outlook API

This is library for easy work with the Outlook Interop API

## Features

- Sending a mail to one or a group of people with or without attachments.
- Reading mails from various client mail folders with or without loading attachments.
- Getting a list of contacts from various address folders.

## Support

Basic information on the Outlook API can be found at: 
https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook?view=outlook-pia

## Usage

### Connection

Add a reference to `OutlookAPI.dll` and import namespace:

```csharp
using OutlookAPI;
```

### Sending a mail to one people

- Without attachments:

```csharp
using OutlookAPI;

//...

bool isOpened = Outlook.OpenApp();
if (isOpened)
{
//...
	
bool isSent = Outlook.SendMail("unknown@domain.com", "Test sending", "This test mail");
if (isSent)
{
//...
}
	
//...
}

//...
``` 

- With attachments:

```csharp
using OutlookAPI;

//...

bool isOpened = Outlook.OpenApp();
if (isOpened)
{
	//...
	
	bool isSent = Outlook.SendMail("unknown@domain.com", "Test sending", "This test mail", new List<string> { "‪C:\\Image.png" });
	if (isSent)
	{
		//...
	}
	
	//...
}

//...
```

### Sending a mail to group people

- Without attachments:

```csharp
using OutlookAPI;

//...

bool isOpened = Outlook.OpenApp();
if (isOpened)
{
	//...
	
	bool isSent = Outlook.SendMail(new List<string> { "unknown1@domain.com", "unknown2@domain.com" }, "Test sending", "This test mail");
	if (isSent)
	{
		//...
	}
	
	//...
}

//...
```

- With attachments:

```csharp
using OutlookAPI;

//...

bool isOpened = Outlook.OpenApp();
if (isOpened)
{
	//...
	
	bool isSent = Outlook.SendMail(new List<string> { "unknown1@domain.com", "unknown2@domain.com" }, "Test sending", "This test mail", new List<string> { "‪C:\\Image.png" });
	if (isSent)
	{
		//...
	}
	
	//...
}

//...
```

### Reading mails without loading attachments

- Without loading attachments:

```csharp
using OutlookAPI;
using OutlookAPI.Models;

//...

bool isOpened = Outlook.OpenApp();
if (isOpened)
{
	//...
	
	List<Mail> mails = OutlookHelper.ReadMails(OutlookHelper.MailFolder.Inbox);
	if (mails != null)
	{
		//...
	}
	
	//...
}

//...
```

- With loading attachments:

```csharp
using OutlookAPI;
using OutlookAPI.Models;

//...

bool isOpened = OpenApp();
if (isOpened)
{
	//...
	
	List<Mail> mails = Outlook.ReadMails(OutlookHelper.MailFolder.Inbox, "‪C:\\");
	if (mails != null)
	{
		//...
	}
	
	//...
}

//...
```

### Getting contacts

```csharp
using OutlookAPI;
using OutlookAPI.Models;

//...

bool isOpened = Outlook.OpenApp();
if (isOpened)
{
	//...
	
	List<Contact> contacts = Outlook.GetContacts(OutlookHelper.AddressBook.GlobalAddressList);
	if (contacts != null)
	{
		//...
	}
	
	//...
}

//...
```