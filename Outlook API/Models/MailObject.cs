using Microsoft.Office.Interop.Outlook;

namespace OutlookAPI.Models
{
    public abstract class MailObject
    {
        #region Methods

        #region Protected internal
        protected internal static bool IsExchangeUser(string typeUser)
        {
            if (typeUser == "EX")
            {
                return true;
            }

            return false;
        }

        protected internal static bool IsCorrectExchangeUser(ExchangeUser user)
        {
            if (user != null)
            {
                return true;
            }

            return false;
        }
        #endregion Protected internal

        #endregion Methods
    }
}