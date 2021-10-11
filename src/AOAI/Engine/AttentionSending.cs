using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using AOAI.Servicing;
using AOAI.Components;

namespace AOAI.Engine
{
    /// <summary>
    /// Functionality for displaying a notification when sending a message
    /// </summary>
    class AttentionSending : FunctionAdd_Ins
    {
        private Outlook.Application _app;

        private List<string> _listHomeDomain;
        private POCO.ConfigAttentionSending _confFunc;
        private bool _isFoundSuperUser = false;

        public override FunctionFeatures FunctionFeature { get; } = FunctionFeatures.AttentionSending;
        public AttentionSending(Outlook.Application application)
        {
            _app = application;
            LoadConfig();
            SubscribingToEvents();
        }

        /// <summary>
        /// Loading the necessary parameters for this component of the add-on from the global configuration
        /// </summary>
        public override void LoadConfig()
        {
            _confFunc = Config.Instance.ConfigAttentionSending ?? new POCO.ConfigAttentionSending();
            _listHomeDomain = Config.Instance.ListHomeDomain ?? new List<string>();
            _isFoundSuperUser = Worker.IsIgnoredFunctionSuperUser(FunctionFeature);
        }

        /// <summary>
        /// Subscribing to Outlook events
        /// </summary>
        private void SubscribingToEvents()
        {
            if (!_confFunc.IsEnableFunction)
                return;
            if (_isFoundSuperUser)
                return;

            _app.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(ConfirmationBeforeSending);
        }

        /// <summary>
        /// The main function is action
        /// Confirmation before sending emails
        /// </summary>
        private void ConfirmationBeforeSending(object Item, ref bool Cancel)
        {
            Outlook.MailItem mailItem = Item as Outlook.MailItem;
            if (mailItem == null || mailItem.Class != Outlook.OlObjectClass.olMail)
                return;

            //If the mail fits, then we perform the action
            if (!MailMatchesForProcessing(mailItem))
                return;

            //We wrap the Type-value in object-value for transfer to another class
            //in order to manipulate exactly this memory area of the type-value object
            Cancel = true;
            List<bool> wrapCancel = new List<bool>() { Cancel };
            AttentionForm form = new AttentionForm(mailItem, wrapCancel, _confFunc);
            form.ShowDialog();
            Cancel = wrapCancel[0];
        }

        /// <summary>
        /// Search for an external address - in the Destination
        /// </summary>
        /// <returns></returns>
        public bool IsExternalMail(Outlook.MailItem mailItem)
        {
            if (mailItem == null || mailItem.Class != Outlook.OlObjectClass.olMail)
                return false;

            List<POCO.MailTransferPerson> POCORecipientList = new List<POCO.MailTransferPerson>();
            Outlook.Recipients recipients = mailItem.Recipients;
            if (recipients.Count < 1 | _listHomeDomain.Count < 1)
                return false;

            foreach (Outlook.Recipient recipient in recipients)
            {
                POCORecipientList.Add(
                    new POCO.MailTransferPerson()
                    {
                        Address = recipient?.Address.ToUpper(),
                        Name = recipient?.Name
                    }
                );
            }
            foreach (POCO.MailTransferPerson r in POCORecipientList)
            {
                if (!_listHomeDomain.Exists(m => r.Address.Contains(m.ToUpper())))
                {
                    //External mail
                    return true;
                }
            }
            //Internal mail
            return false;
        }

        /// <summary>
        /// Checking the mail for compliance with the conditions of the specified configuration 
        /// for further processing by the add-on functions
        /// </summary>
        public override bool MailMatchesForProcessing(Outlook.MailItem mailItem)
        {
            bool flag = false;
            //Have external users been found
            if (IsExternalMail(mailItem))
            {
                //Is it necessary to process attachments
                if (!_confFunc.IsEnableCheckAttachment)
                    return flag = true;
                return flag = (Worker.IsWithAttachment(mailItem)) ? true : false;
            }
            return flag;
        }
    }
}
