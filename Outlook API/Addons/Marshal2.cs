using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;

namespace OutlookAPI.Addons
{
    /// <summary>
    /// An addon to the Marshal class that adds a method for 
    /// retrieving the running instance of the specified object from the table of running objects
    /// </summary>
    public static class Marshal2
    {
        #region Constants

        #region Private
        private const string OLEAUT32 = "oleaut32.dll";
        private const string OLE32 = "ole32.dll";
        #endregion Private

        #endregion Constants

        #region Methods 

        #region Private
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [SecurityCritical]
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [SecurityCritical]
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        [DllImport(OLEAUT32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [SecurityCritical]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out object ppunk);
        #endregion Priate

        #region Public
        /// <summary>
        /// Retrieves the running instance of the specified object from the table of running objects (ROT).
        /// </summary>
        /// <param name="progID">
        /// The program identifier (ProgID) of the requested object.
        /// </param>
        /// <returns>
        /// The requested object; otherwise null. This object can be cast to any COM interface supported by this object.
        /// </returns>
        [SecurityCritical]
        public static object GetActiveObject(string progID)
        {
            Guid clsid;

            try
            {
                CLSIDFromProgIDEx(progID, out clsid);
            }
            catch (Exception)
            {
                CLSIDFromProgID(progID, out clsid);
            }

            GetActiveObject(ref clsid, IntPtr.Zero, out object obj);

            return obj;
        }
        #endregion Public

        #endregion Methods
    }
}