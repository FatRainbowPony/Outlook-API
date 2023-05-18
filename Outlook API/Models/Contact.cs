using Microsoft.Office.Interop.Outlook;

namespace OutlookAPI.Models
{
	public class Contact : MailObject
    {
        #region Properties

        #region Public
        public int Id { get; set; }

        public string? Name { get; set; }

        public string? JobTitle { get; set; }

        public string? Department { get; set; }

        public string? Email { get; set; }
        #endregion Public

        #endregion Properties

        #region Constructors

        #region Public
        public Contact()
        {

        }

        public Contact(int id, ExchangeUser userData)
        {
            if (userData != null)
            {
                Id = id;
                Name = userData.Name;
                JobTitle = userData.JobTitle;
                Department = userData.Department;
                Email = userData.PrimarySmtpAddress;
            }
        }
        public Contact(int id, AddressEntry userData)
        {
            if (userData != null)
            {
                Id = id;
                Name = userData.Name;
                Email = userData.Address;
            }
        }

        public Contact(int id, ContactItem userData)
        {
            if (userData != null)
            {
                Id = id;
                Name = userData.FullName;
                JobTitle = userData.JobTitle;
                Department = userData.Department;
                Email = userData.Email1Address;
            }
        }
        #endregion Public

        #endregion Constructors
    }
}