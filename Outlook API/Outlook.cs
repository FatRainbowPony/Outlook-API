using System.Diagnostics;
using System.Runtime.InteropServices;
using OutlookAPI.Models;
using Microsoft.Office.Interop.Outlook;
using OutlookAPI.Addons;

namespace OutlookAPI
{
    /// <summary>
    /// Provides a collection of methods for interacting with the Outlook application
    /// </summary>
    public static class Outlook
    {
        #region Constants

        #region Private
        private const string PROG_ID = "Outlook.Application";
        private const string EXE_NAME = "outlook";
        private const string DEFAULT_NAMESPACE = "MAPI";
        #endregion Private

        #endregion Constants

        #region Enums

        #region Public
        public enum MailFolder
        {
            Inbox,
            Sent,
            Deleted,
            Junk
        }

        public enum AddressBook
        {
            Contacts,
            GlobalAddressList
        }
        #endregion Public

        #endregion Enums

        #region Fields

        #region Private
        private static bool isFirstLaunch;
        private static Application? app;
        #endregion Private

        #endregion Fields

        #region Methods

        #region Private

        #region Method for closing application
        /// <summary>
        /// Closes the Outlook application
        /// </summary>
        private static void CloseApp()
        {
            if (app != null)
            {
                app.Quit();
                ReleaseObj(app);
            }

            static void ReleaseObj(object? obj)
            {
                if (obj != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(obj);
                    }
                    finally
                    {
                        obj = null;
                    }
                }
            }
        }
        #endregion Method for closing application

        #endregion Private

        #region Public

        #region Method for opening application
        /// <summary>
        /// Opens/connects to the application
        /// </summary>
        /// <returns>
        /// true - if the opening/connection is successful, false - if the opening/connection is not successful
        /// </returns>
        public static bool OpenApp()
        {
            if (Type.GetTypeFromProgID(PROG_ID) != null)
            {
                if (Process.GetProcessesByName(EXE_NAME).Length == 0)
                {
                    isFirstLaunch = true;
                    app = new Application();
                }
                else
                {
                    isFirstLaunch = false;
                    app = (Application)Marshal2.GetActiveObject(PROG_ID);
                }

                return true;
            }

            return false;
        }
        #endregion Method for opening application

        #region Methods for sending mail
        /// <summary>
        /// Sends an email to one person without attachments.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="receiver">
        /// Receiver's email address
        /// </param>
        /// <param name="subject">
        /// Subject of the mail
        /// </param>
        /// <param name="body">
        /// The body of the mail
        /// </param>
        /// <returns>
        /// true - if the email has been sent, false - if the email has not been sent.
        /// </returns>
        public static bool SendMail(string receiver, string subject, string body)
        {
            if (app != null)
            {
                if (!string.IsNullOrEmpty(receiver) && !string.IsNullOrWhiteSpace(receiver))
                {
                    var mail = (MailItem)app.CreateItem(OlItemType.olMailItem);
                    mail.To = receiver;
                    mail.Subject = subject;
                    mail.Body = body;
                    mail.Importance = OlImportance.olImportanceNormal;

                    try
                    {
                        mail.Send();

                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }

                if (isFirstLaunch)
                {
                    CloseApp();
                }
            }

            return false;
        }

        /// <summary>
        /// Sends an email to group persons without attachments.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="receivers">
        /// List of receivers emails addresses
        /// </param>
        /// <param name="subject">
        /// Subject of the mail
        /// </param>
        /// <param name="body">
        /// The body of the mail
        /// </param>
        /// <returns>
        /// true - if the email has been sent, false - if the email has not been sent.
        /// </returns>
        public static bool SendMail(List<string> receivers, string subject, string body)
        {
            return SendMail(string.Join(";", receivers), subject, body);
        }

        /// <summary>
        /// Sends an email to one person with attachments.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="receiver">
        /// Receiver's email address
        /// </param>
        /// <param name="subject">
        /// Subject of the mail
        /// </param>
        /// <param name="body">
        /// The body of the mail
        /// </param>
        /// <param name="attachments">
        /// List of mail attachments
        /// </param>
        /// <returns>
        /// true - if the email has been sent, false - if the email has not been sent.
        /// </returns>
        public static bool SendMail(string receiver, string subject, string body, List<string> attachments)
        {
            if (app != null)
            {
                if (attachments == null || attachments.Count <= 0)
                {
                    return SendMail(receiver, subject, body);
                }
                else
                {
                    if (!string.IsNullOrEmpty(receiver) && !string.IsNullOrWhiteSpace(receiver))
                    {
                        var mail = (MailItem)app.CreateItem(OlItemType.olMailItem);
                        mail.To = receiver;
                        mail.Subject = subject;
                        mail.Body = body;
                        mail.Importance = OlImportance.olImportanceNormal;

                        foreach (var attachment in attachments)
                        {
                            mail.Attachments.Add(attachment, (int)OlAttachmentType.olByValue, mail.Body.Length + 1, Path.GetFileNameWithoutExtension(attachment));
                        }

                        try
                        {
                            mail.Send();
                            
                            return true;
                        }
                        catch
                        {
                            return false;
                        }
                    }
                }

                if (isFirstLaunch)
                {
                    CloseApp();
                }
            }

            return false;
        }

        /// <summary>
        /// Sends an email to group persons with attachments.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="receivers">
        /// List of receivers emails addresses
        /// </param>
        /// <param name="subject">
        /// Subject of the mail
        /// </param>
        /// <param name="body">
        /// The body of the mail
        /// </param>
        /// <param name="attachments">
        /// List of mail attachments
        /// </param>
        /// <returns>
        /// true - if the email has been sent, false - if the email has not been sent.
        /// </returns>
        public static bool SendMail(List<string> receivers, string subject, string body, List<string> attachments)
        {
            return SendMail(string.Join(";", receivers), subject, body, attachments);
        }
        #endregion Methods for sending mail

        #region Methods for reading mails
        /// <summary>
        /// Reads mails from a mail folder. 
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="mailFolder">
        /// The mail folder from which mails are read
        /// </param>
        /// <returns>
        /// Mail list
        /// </returns>
        public static List<Mail>? ReadMails(MailFolder mailFolder)
        {
            List<Mail>? mails = null;

            if (app != null)
            {
                var nameSpace = app.GetNamespace(DEFAULT_NAMESPACE);
                if (nameSpace != null)
                {
                    MAPIFolder? folder = null;
                    switch (mailFolder)
                    {
                        case MailFolder.Inbox:
                            folder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                            break;

                        case MailFolder.Sent:
                            folder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
                            break;

                        case MailFolder.Deleted:
                            folder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
                            break;

                        case MailFolder.Junk:
                            folder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderJunk);
                            break;
                    }

                    if (folder != null && folder.Items.Count > 0)
                    {
                        mails = new List<Mail>();

                        foreach (var folderItem in folder.Items)
                        {
                            if (folderItem != null && folderItem is MailItem mailItem)
                            {
                                mails.Add(new Mail(mailItem));
                            }
                        }
                    }
                }
            }

            return mails;
        }

        /// <summary>
        /// Reads mails from a mail folder and saves attachments to the specified folder.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="mailFolder">
        /// The mail folder from which mails are read
        /// </param>
        /// <param name="pathToDirForSavingAttachments">
        /// The path to the directory to save the attachment
        /// </param>
        /// <returns>
        /// Mail list
        /// </returns>
        public static List<Mail>? ReadMails(MailFolder mailFolder, string pathToDirForSavingAttachments)
        {
            List<Mail>? mails = null;

            if (app != null)
            {
                var nameSpace = app.GetNamespace(DEFAULT_NAMESPACE);
                if (nameSpace != null)
                {
                    MAPIFolder? folder = null;
                    switch (mailFolder)
                    {
                        case MailFolder.Inbox:
                            folder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                            break;

                        case MailFolder.Sent:
                            folder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderSentMail);
                            break;

                        case MailFolder.Deleted:
                            folder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
                            break;

                        case MailFolder.Junk:
                            folder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderJunk);
                            break;
                    }

                    if (folder != null && folder.Items.Count > 0)
                    {
                        mails = new List<Mail>();

                        foreach (var folderItem in folder.Items)
                        {
                            if (folderItem != null && folderItem is MailItem mailItem)
                            {
                                mails.Add(new Mail(mailItem));

                                if (Directory.Exists(pathToDirForSavingAttachments) &&
                                    mailItem.Attachments != null && mailItem.Attachments.Count > 0)
                                {
                                    foreach (Attachment attachment in mailItem.Attachments)
                                    {
                                        attachment.SaveAsFile(Path.Combine(pathToDirForSavingAttachments, attachment.FileName));
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return mails;
        }
        #endregion Methosd for reading mails

        #region Method for getting contacts
        /// <summary>
        /// Gets the contacts list from the address book.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="addressBook">
        /// Sets the target address book for receiving contacts 
        /// (if Contacts, a list will be received from the contact book; 
        /// if GlobalAddressList, a list will be received from the global address book)
        /// </param>
        /// <returns>
        /// Contact list
        /// </returns>
        public static List<Contact>? GetContacts(AddressBook addressBook)
        {
            List<Contact>? contacts = null;

            if (app != null)
            {
                object addressList;

                switch (addressBook)
                {
                    case AddressBook.Contacts:
                        addressList = app.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts).Items;
                        if (addressList != null)
                        {
                            contacts = new List<Contact>();

                            for (int i = 1; i <= ((Items)addressList).Count; i++)
                            {
                                contacts.Add(new Contact(i, (ContactItem)((Items)addressList)[i]));
                            }
                        }
                        break;

                    case AddressBook.GlobalAddressList:
                        addressList = app.Session.GetGlobalAddressList();
                        if (addressList != null)
                        {
                            var addressEntries = ((AddressList)addressList).AddressEntries;
                            if (addressEntries != null && addressEntries.Count > 0)
                            {
                                contacts = new List<Contact>();

                                for (int i = 1; i <= addressEntries.Count; i++)
                                {
                                    if (Contact.IsExchangeUser(addressEntries[i].Type) && Contact.IsCorrectExchangeUser(addressEntries[i].GetExchangeUser()))
                                    {
                                        contacts.Add(new Contact(i, addressEntries[i].GetExchangeUser()));
                                    }
                                    else
                                    {
                                        contacts.Add(new Contact(i, addressEntries[i]));
                                    }
                                }
                            }
                        }
                        break;
                }

                if (isFirstLaunch)
                {
                    CloseApp();
                }
            }

            return contacts;
        }
        #endregion Method for getting contacts

        #region Method for getting a list of contact names
        /// <summary>
        /// Gets the name list from the address book.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="addressBook">
        /// Sets the target address book for receiving contacts 
        /// (if Contacts, a list will be received from the contact book; 
        /// if GlobalAddressList, a list will be received from the global address book)
        /// </param>
        /// <returns>
        /// Name list
        /// </returns>
        public static List<string>? GetNames(AddressBook addressBook)
        {
            List<string>? names = null;

            if (app != null)
            {
                var contacts = GetContacts(addressBook);
                if (contacts != null)
                {
                    names = new List<string>();
                    foreach (var contact in contacts)
                    {
                        if (contact.Name != null)
                        {
                            names.Add(contact.Name);
                        }
                    }
                }

                if (isFirstLaunch)
                {
                    CloseApp();
                }
            }

            return names;
        }
        #endregion Method for getting a list of contact names

        #region Method for getting a list of contact job titles
        /// <summary>
        /// Gets the job title list from the address book.
        /// First need to call the method OpenApp()
        /// <param name="addressBook">
        /// Sets the target address book for receiving contacts 
        /// (if Contacts, a list will be received from the contact book; 
        /// if GlobalAddressList, a list will be received from the global address book)
        /// </param>
        /// </summary>
        /// <returns>
        /// Job title list
        /// </returns>
        public static List<string>? GetJobTitles(AddressBook addressBook)
        {
            List<string>? jobTitles = null;

            if (app != null)
            {
                var contacts = GetContacts(addressBook);
                if (contacts != null)
                {
                    jobTitles = new List<string>();
                    foreach (var contact in contacts)
                    {
                        if (contact.JobTitle != null)
                        {
                            jobTitles.Add(contact.JobTitle);
                        }
                    }
                }

                if (isFirstLaunch)
                {
                    CloseApp();
                }
            }

            return jobTitles;
        }
        #endregion Method for getting a list of contact job titles

        #region Method for getting a list of contact departments
        /// <summary>
        /// Gets the department list from the address book.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="addressBook">
        /// Sets the target address book for receiving contacts 
        /// (if Contacts, a list will be received from the contact book; 
        /// if GlobalAddressList, a list will be received from the global address book)
        /// </param>
        /// <returns>
        /// Department list
        /// </returns>
        public static List<string>? GetDepartments(AddressBook addressBook)
        {
            List<string>? departments = null;

            if (app != null)
            {
                var contacts = GetContacts(addressBook);
                if (contacts != null)
                {
                    departments = new List<string>();

                    foreach (var contact in contacts)
                    {
                        if (contact.Department != null)
                        {
                            departments.Add(contact.Department);
                        }
                    }
                }

                if (isFirstLaunch)
                {
                    CloseApp();
                }
            }

            return departments;
        }
        #endregion Method for getting a list of contact departments

        #region Method for getting a list of contact e-mails
        /// <summary>
        /// Gets the e-mail list from the address book.
        /// First need to call the method OpenApp()
        /// </summary>
        /// <param name="addressBook">
        /// Sets the target address book for receiving contacts 
        /// (if Contacts, a list will be received from the contact book; 
        /// if GlobalAddressList, a list will be received from the global address book)
        /// </param>
        /// <returns>
        /// E-mail list
        /// </returns>
        public static List<string>? GetEmails(AddressBook addressBook)
        {
            List<string>? emails = null;

            if (app != null)
            {
                var contacts = GetContacts(addressBook);
                if (contacts != null)
                {
                    emails = new List<string>();

                    foreach (var contact in contacts)
                    {
                        if (contact.Email != null)
                        {
                            emails.Add(contact.Email);
                        }
                    }
                }

                if (isFirstLaunch)
                {
                    CloseApp();
                }
            }

            return emails;
        }
        #endregion Method for getting a list of contact e-mails

        #endregion Public

        #endregion Methods
    }
}