using System;
using System.DirectoryServices;

namespace AOAI.Components
{
    /// <summary>
    /// Assumes that the user is working in Active Directory. 
    /// The extension searches for a "contact object" named "AOAI*" in AD, 
    /// reads the XML configuration from the "houseIdentifier" attribute
    /// </summary>
    class ExternalStorageConfigAD
    {
        private string _domainNameStr;
        private string _contactNameStr;
        public ExternalStorageConfigAD(string domainNameStr, string contactNameStr)
        {
            _domainNameStr = domainNameStr;
            _contactNameStr = contactNameStr;
        }
        public string GetData()
        {
            string resultString = null;
            try
            {
                DirectoryEntry de = new DirectoryEntry($"LDAP://{_domainNameStr}");
                using (DirectorySearcher deSearcher = new DirectorySearcher(de))
                {
                    deSearcher.Filter = $"(&(objectCategory=person)(cn={_contactNameStr}))";
                    deSearcher.SearchScope = SearchScope.Subtree;
                    deSearcher.ClientTimeout = new TimeSpan(0, 0, 4);

                    SearchResult searchEntry = deSearcher.FindOne();
                    if (searchEntry != null)
                    {
                        DirectoryEntry retriviedDe = searchEntry.GetDirectoryEntry();
                        object value = retriviedDe.Properties["houseIdentifier"].Value;

                        if (value is string valueString)
                            resultString = valueString;
                        if (value is object[] valueStringArray)
                            resultString = (string)valueStringArray[0];

                        retriviedDe.Dispose();
                    }
                    de.Dispose();
                }
            }
            catch
            {
                resultString = null;
            }
            return resultString;
        }
    }
}
