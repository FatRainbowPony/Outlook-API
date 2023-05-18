using Microsoft.Office.Interop.Outlook;

namespace OutlookAPI.Models
{
	public class Mail : MailObject
    {
        #region Structures

        #region Public
        public struct SenderInfo
        {
            public string name;
            public string email;
        }

        public struct ReceiversInfo
        {
            public string name;
            public string email;
        }
        #endregion Public

        #endregion Structures

        #region Properties

        #region Public
        public DateTime? CreationDate { get; set; }

        public SenderInfo? Sender { get; set; }

        public List<ReceiversInfo>? Receivers { get; set; }

        public string? Subject { get; set; }

        public List<string>? Attachments { get; set; }

        public string? Body { get; set; }
        #endregion Public

        #endregion Properties

        #region Constructors

        #region Public
        public Mail()
        {

        }

        public Mail(MailItem mail)
        {
            if (mail != null)
            {
                CreationDate = mail.CreationTime;

                if (mail.Sender != null)
                {
                    if (IsExchangeUser(mail.Sender.Type) && IsCorrectExchangeUser(mail.Sender.GetExchangeUser()))
                    {
                        Sender = new SenderInfo { name = mail.SenderName, email = mail.Sender.GetExchangeUser().PrimarySmtpAddress };
                    }
                    else
                    {
                        Sender = new SenderInfo { name = mail.SenderName, email = mail.Sender.Address };
                    }
                }

                if (mail.Recipients != null && mail.Recipients.Count > 0)
                {
                    Receivers = new List<ReceiversInfo>();

                    foreach (Recipient recipient in mail.Recipients)
                    {
                        if (IsExchangeUser(recipient.AddressEntry.Type) && IsCorrectExchangeUser(recipient.AddressEntry.GetExchangeUser()))
                        {
                            Receivers.Add(new ReceiversInfo { name = recipient.Name, email = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress });
                        }
                        else
                        {
                            Receivers.Add(new ReceiversInfo { name = recipient.Name, email = recipient.Address });
                        }
                    }
                }

                Subject = mail.Subject;

                if (mail.Attachments != null && mail.Attachments.Count > 0)
                {
                    Attachments = new List<string>();

                    foreach (Attachment attachment in mail.Attachments)
                    {
                        Attachments.Add(attachment.FileName);
                    }
                }

                Body = mail.Body;
            }
        }
        #endregion Public

        #endregion Constructors
    }
}