namespace AOAI.Components
{
    /// <summary>
    /// Local instance of the XML configuration
    /// </summary>
    class LocalStorageConfig
    {
        public string GetData()
        {
            string resultString = "<?xml version=\"1.0\" encoding=\"utf-16\"?>" +
                                "<Config xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://www.default.www\">" +
                                "  <ListHomeDomain>" +
                                "    <DomainObject>o=DOMAIN1TEST</DomainObject>" +
                                "    <DomainObject>@domain.1test</DomainObject>" +
                                "    <DomainObject>@anotherdomain.1test</DomainObject>" +
                                "  </ListHomeDomain>" +
                                "  <ListIgnoredUsers>" +
                                "    <SuperUser>" +
                                "      <Name>DOMAIN1TEST\\_KUL.teSt</Name>" +
                                "      <ignoredFunctions>" +
                                "        <FunctionFeatures>AttentionSending</FunctionFeatures>" +
                                "        <FunctionFeatures>MarkingMail</FunctionFeatures>" +
                                "      </ignoredFunctions>" +
                                "    </SuperUser>" +
                                "  </ListIgnoredUsers>" +
                                "  <ConfigAttentionWhenSending>" +
                                "    <IsEnableFunction>true</IsEnableFunction>" +
                                "    <IsEnableCheckAttachment>false</IsEnableCheckAttachment>" +
                                "    <FormTextLabel>ATTENTION!</FormTextLabel>" +
                                "    <BtnSend>Send</BtnSend>" +
                                "    <BtnCancel>Cancel</BtnCancel>" +
                                "    <LabelInformation>" +
                                "      <LabelTextObject>" +
                                "        <Text>You are sending an email to an external recipient!\n</Text>" +
                                "        <Color>Black</Color>" +
                                "      </LabelTextObject>" +
                                "      <LabelTextObject>" +
                                "        <Text>Make sure that the email meets all the company's standards and policy before sending.</Text>" +
                                "        <Color>Red</Color>" +
                                "      </LabelTextObject>" +
                                "    </LabelInformation>" +
                                "  </ConfigAttentionWhenSending>" +
                                "  <ConfigMarkingExternalEmails>" +
                                "    <IsEnableFunction>true</IsEnableFunction>" +
                                "    <IsEnableHandlingNewMail>true</IsEnableHandlingNewMail>" +
                                "    <IsEnableHandlingFolder>true</IsEnableHandlingFolder>" +
                                "    <IsEnableHandlingUI>true</IsEnableHandlingUI>" +
                                "    <IsEnableMarkedExternalRecipientsOnSentMail>true</IsEnableMarkedExternalRecipientsOnSentMail>" +
                                "    <IsEnableAddititionalyHandlingOnlyUnread>false</IsEnableAddititionalyHandlingOnlyUnread>" +
                                "    <MailCategoryLabel>EXTERNAL EMAIL</MailCategoryLabel>" +
                                "    <MailCategoryColor>11</MailCategoryColor>" +
                                "  </ConfigMarkingExternalEmails>" +
                                "</Config>";
            return resultString;
        }
    }
}
