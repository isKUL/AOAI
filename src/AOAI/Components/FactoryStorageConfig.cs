using System;
using System.Net.NetworkInformation;
using AOAI.Servicing;

namespace AOAI.Components
{
    /// <summary>
    /// Getting the XML configuration from the specified storage
    /// </summary>
    class FactoryStorageConfig
    {
        StorageConfig _exStorConf;
        public FactoryStorageConfig(StorageConfig exStorConf)
        {
            _exStorConf = exStorConf;
        }
        public String GetConfig()
        {
            String xmlConfStr = null;
            switch (_exStorConf)
            {
                case StorageConfig.LocalStorage:
                    LocalStorageConfig localStor = new LocalStorageConfig();
                    xmlConfStr = localStor.GetData();
                    break;

                case StorageConfig.ActiveDirectory:
                    string servingDomain = IPGlobalProperties.GetIPGlobalProperties().DomainName;
                    if (!String.IsNullOrEmpty(servingDomain))
                    {
                        ExternalStorageConfigAD extStorAD = new ExternalStorageConfigAD(servingDomain, "*AOAI*");
                        xmlConfStr = extStorAD.GetData();
                    }
                    break;
            }
            return xmlConfStr;
        }

    }
}
