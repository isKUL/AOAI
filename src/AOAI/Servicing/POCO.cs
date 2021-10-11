using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace AOAI.Servicing
{
    /// <summary>
    /// Class-container of auxiliary elements
    /// </summary>
    [Serializable]
    public class POCO
    {
        public class MailTransferPerson
        {
            public string Address { get; set; }
            public string Name { get; set; }
        }

        /// <summary>
        /// Attachment properties of mail
        /// </summary>
        public class Attachment
        {
            public string FileName { get; set; }
            public int Size { get; set; }
        }

        /// <summary>
        /// User who ignores the functionality
        /// </summary>
        [Serializable]
        public class SuperUser
        {
            public string Name { get; set; }
            public List<FunctionFeatures> ignoredFunctions { get; set; } = new List<FunctionFeatures>();
        }

        /// <summary>
        /// Text element, of the notification display functionality, form
        /// </summary>
        [Serializable]
        public class LabelTextObject
        {
            public string Text { get; set; }
            public string Color { get; set; }
        }

        /// <summary>
        /// Configuration of the displaying a notification functionality
        /// </summary>
        [Serializable]
        public class ConfigAttentionSending
        {
            [XmlElement("IsEnableFunction")]
            public bool IsEnableFunction { get; set; } = false;

            [XmlElement("IsEnableCheckAttachment")]
            public bool IsEnableCheckAttachment { get; set; }

            [XmlElement("FormTextLabel")]
            public string FormTextLabel { get; set; }

            [XmlElement("BtnSend")]
            public string BtnSendLabel { get; set; }

            [XmlElement("BtnCancel")]
            public string BtnCancelLabel { get; set; }

            [XmlElement("CheckBoxAccept")]
            public string CheckBoxAcceptLabel { get; set; }

            [XmlArray]
            [XmlArrayItem("LabelTextObject")]
            public List<POCO.LabelTextObject> LabelInformation { get; set; } = new List<LabelTextObject>();
        }

        /// <summary>
        /// Configuration of the mail marking functionality
        /// </summary>
        [Serializable]
        public class ConfigMarkingMail
        {
            [XmlElement("IsEnableFunction")]
            public bool IsEnableFunction { get; set; } = false;

            [XmlElement("IsEnableHandlingNewMail")]
            public bool IsEnableHandlingNewMail { get; set; } = true;

            [XmlElement("IsEnableHandlingFolder")]
            public bool IsEnableHandlingFolder { get; set; }

            [XmlElement("IsEnableHandlingUI")]
            public bool IsEnableHandlingUI { get; set; }

            [XmlElement("IsEnableMarkedExternalRecipientsOnSentMail")]
            public bool IsEnableMarkedExternalRecipientsOnSentMail { get; set; }

            [XmlElement("IsEnableAddititionalyHandlingOnlyUnread")]
            public bool IsEnableAddititionalyHandlingOnlyUnread { get; set; }

            [XmlElement("MailCategoryLabel")]
            public string MailCategoryLabel { get; set; } = "EXTERNAL EMAIL";

            [XmlElement("MailCategoryColor")]
            public int MailCategoryColor { get; set; } = 11; //Steel
        }
    }
}
