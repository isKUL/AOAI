using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.Xml;
using System.IO;
using AOAI.Servicing;

namespace AOAI.Components
{
    /// <summary>
    /// A class for representing the global configuration of this extension. 
    /// The configuration is loaded dynamically.
    /// </summary>
    [Serializable]
    [XmlRoot("Config")]
    public sealed class Config
    {
        private static volatile Config _instance;
        private static readonly object _sync = new object();

        private List<string> _listHomeDomain = new List<string>();
        private List<POCO.SuperUser> _listIgnoredUsers = new List<POCO.SuperUser>();

        private static Config XmlDeserialize(String xmlConfStr)
        {
            Config classConfig = null;
            if (String.IsNullOrEmpty(xmlConfStr))
                return null;
            XmlRootAttribute xmlRoot = new XmlRootAttribute("Config");
            xmlRoot.IsNullable = true;
            xmlRoot.Namespace = "http://www.default.www";
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Config), xmlRoot);
            using (StringReader textReader = new StringReader(xmlConfStr))
            {
                try
                {
                    classConfig = (Config)xmlSerializer.Deserialize(textReader);
                }
                catch
                {
                    classConfig = null;
                }
            }
            return classConfig;
        }

        private static string XmlSerialize()
        {
            string xmlConfStr = null;
            XmlRootAttribute xmlRoot = new XmlRootAttribute("Config");
            xmlRoot.IsNullable = true;
            xmlRoot.Namespace = "http://www.default.www";
            XmlSerializer mySerializer = new XmlSerializer(typeof(Config), xmlRoot);
            using (StringWriter textWriter = new StringWriter())
            {
                mySerializer.Serialize(textWriter, Config.Instance);
                xmlConfStr = textWriter.ToString();
            }
            return xmlConfStr;
        }

        public static Config Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_sync)
                    {
                        if (_instance == null)
                        {
                            _instance = new Config();
                        }
                    }
                }
                return _instance;
            }
            set { if (value != null) _instance = value; }
        }

        [XmlArray]
        [XmlArrayItem("DomainObject")]
        public List<string> ListHomeDomain
        {
            get { return _listHomeDomain; }
            set { if (value != null) _listHomeDomain = value; }
        }

        [XmlArray]
        [XmlArrayItem("SuperUser")]
        public List<POCO.SuperUser> ListIgnoredUsers
        {
            get { return _listIgnoredUsers; }
            set { if (value != null) _listIgnoredUsers = value; }
        }

        [XmlElement("ConfigAttentionWhenSending")]
        public POCO.ConfigAttentionSending ConfigAttentionSending { get; set; } = new POCO.ConfigAttentionSending();

        [XmlElement("ConfigMarkingExternalEmails")]
        public POCO.ConfigMarkingMail ConfigMarkingMail { get; set; } = new POCO.ConfigMarkingMail();

        //For debugging
        public static void SaveConfig()
        {
            var xmlConfStr = XmlSerialize();
            File.WriteAllText(Path.Combine(Environment.CurrentDirectory, "Config.xml"), xmlConfStr);
        }
        public static void LoadConfig()
        {
            FactoryStorageConfig factoryStorConf = null;
            String xmlConfStr = null;
            Config confFromXml = null;
            Dictionary<StorageConfig, String> xmlConfDict = new Dictionary<StorageConfig, string>();
            Dictionary<StorageConfig, Config> confDict = new Dictionary<StorageConfig, Config>();
            String[] collectionStorages = new string[] {
                                                         StorageConfig.LocalStorage.ToString(),
                                                         StorageConfig.ActiveDirectory.ToString()
                                                       };

            foreach (string storageXmlConf in collectionStorages)
            {
                StorageConfig elemEnum = (StorageConfig)Enum.Parse(typeof(StorageConfig), storageXmlConf);
                factoryStorConf = new FactoryStorageConfig(elemEnum);
                try
                {
                    xmlConfStr = factoryStorConf.GetConfig();
                }
                catch
                {
                    xmlConfStr = null;
                }
                if (String.IsNullOrEmpty(xmlConfStr))
                    continue;
                xmlConfDict.Add(elemEnum, xmlConfStr);
            }

            foreach (KeyValuePair<StorageConfig, String> elemStorXmlConf in xmlConfDict)
            {
                try
                {
                    confFromXml = XmlDeserialize(elemStorXmlConf.Value);
                    if (confFromXml == null)
                        continue;
                    confDict.Add(elemStorXmlConf.Key, confFromXml);
                }
                catch { }
            }

            if (confDict.ContainsKey(StorageConfig.ActiveDirectory))
            {
                Config.Instance = confDict[StorageConfig.ActiveDirectory];
                return;
            }
            if (confDict.ContainsKey(StorageConfig.LocalStorage))
            {
                Config.Instance = confDict[StorageConfig.LocalStorage];
                return;
            }
        }
    }
}
