using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using AOAI.Servicing;
using AOAI.Components;
using Microsoft.Win32;

namespace AOAI.Engine
{
    /// <summary>
    /// Functionality for marking incoming and sending emails
    /// </summary>
    class MarkingMail : FunctionAdd_Ins
    {
        private delegate bool IsExternal(Outlook.MailItem m);
        IsExternal IsExternalMail;

        private Outlook.Application _app;
        private Outlook.Explorer _explorer;
        private List<Outlook.Items> _itemsList;

        private List<string> _listHomeDomain;
        private POCO.ConfigMarkingMail _confFunc;
        private char categorySeparator;
        private bool _isFoundSuperUser;

        public override FunctionFeatures FunctionFeature { get; } = FunctionFeatures.MarkingMail;
        public MarkingMail(Outlook.Application application)
        {
            _app = application;
            _itemsList = new List<Outlook.Items>();
            LoadConfig();
            VerifyCategoryOnOutlook(_confFunc.MailCategoryLabel, (Outlook.OlCategoryColor)_confFunc.MailCategoryColor);
            SubscribingToEvents();
        }

        /// <summary>
        /// Loading the necessary parameters for this component of the add-on from the global configuration
        /// </summary>
        public override void LoadConfig()
        {
            _confFunc = Config.Instance.ConfigMarkingMail ?? new POCO.ConfigMarkingMail();
            _listHomeDomain = Config.Instance.ListHomeDomain ?? new List<string>();
            _isFoundSuperUser = Worker.IsIgnoredFunctionSuperUser(FunctionFeature);

            //Get an international separator symbol in the user's environment
            try
            {
                String tmpSeparator = (String)(Registry.CurrentUser.OpenSubKey("Control Panel\\International")?.GetValue("sList"));
                if (!String.IsNullOrEmpty(tmpSeparator) && tmpSeparator.Length == 1)
                    categorySeparator = tmpSeparator[0];
                else
                    categorySeparator = ';';
            }
            catch
            {
                categorySeparator = ';';
            }
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

            //New messages
            if (_confFunc.IsEnableHandlingNewMail)
                _app.NewMailEx += new Outlook.ApplicationEvents_11_NewMailExEventHandler(NewMail);
            //Clicks in explorer
            if (_confFunc.IsEnableHandlingUI)
                ConnectUI();
            //Actions in directories
            if (_confFunc.IsEnableHandlingFolder)
                ConnectMailFolder();
        }

        /// <summary>
        /// Handle events when new incoming emails arrive
        /// </summary>
        private void NewMail(string EntryIDCollection)
        {
            if (String.IsNullOrEmpty(EntryIDCollection))
                return;

            Outlook.MailItem mailItem = null;
            try
            {
                mailItem = _app.Session.GetItemFromID(EntryIDCollection) as Outlook.MailItem;
            }
            catch
            {
                return;
            }

            CatchMail(mailItem);
        }

        /// <summary>
        /// Handle mail directory events for adding/changing emails
        /// </summary>
        private void ConnectMailFolder()
        {
            AnalysisFolder(_app.Session.Folders);

            foreach (Outlook.Items mailItems in _itemsList)
            {
                try
                {
                    mailItems.ItemChange += new Outlook.ItemsEvents_ItemChangeEventHandler(CatchMail);
                    mailItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(CatchMail);
                }
                catch { }
            }
        }

        //We go through all the directories inside the mailbox
        //and assign the collections of emails to the list
        private void AnalysisFolder(Outlook.Folders folders)
        {
            Outlook.Folder folder = folders.GetFirst() as Outlook.Folder;
            do
            {
                if (folder != null)
                {
                    //We do not process the directory of Outbox/Drafts emails - can catch an exception
                    if (
                        _app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox).FolderPath == folder.FolderPath
                        || _app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts).FolderPath == folder.FolderPath
                        )
                    {
                        folder = folders.GetNext() as Outlook.Folder;
                        continue;
                    }

                    //If Outlook is disabled, there will be exceptions when trying to connect to the synchronization log directories
                    try
                    {
                        _itemsList.Add(folder.Items);

                        //If the directory has subdirectories, then iterate through them
                        if (folder.Folders.Count >= 1)
                            AnalysisFolder(folder.Folders);
                    }
                    catch { }
                }

                folder = folders.GetNext() as Outlook.Folder;
            }
            while (folder != null);
        }

        /// <summary>
        /// Handle user interface events.
        /// In particular, clicking in Outlook Explorer
        /// </summary>
        private void ConnectUI()
        {
            try
            {
                _explorer = _app.ActiveExplorer();
                _explorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(ViewSelectedItem);
            }
            catch { }
        }

        private void ViewSelectedItem()
        {
            var a = _app.ActiveExplorer();
            Outlook.Selection selection = _app.ActiveExplorer().Selection;
            for (int i = 1; i <= selection.Count; i++)
            {
                Outlook.MailItem mailItem = selection[i] as Outlook.MailItem;
                CatchMail(mailItem);
            }
        }

        /// <summary>
        /// Input function for evaluating a email before coloring it
        /// </summary>
        /// <param name="item">The object of the Outlook message</param>
        public void CatchMail(object item)
        {
            Outlook.MailItem mailItem;
            Outlook.Folder folder;
            try
            {
                mailItem = item as Outlook.MailItem;
                if (mailItem == null || mailItem.Class != Outlook.OlObjectClass.olMail)
                    return;

                folder = mailItem.Parent as Outlook.Folder;
                if (folder == null)
                    return;
            }
            catch { return; }

            if (!MailMatchesForProcessing(mailItem))
                return;

            //In the catalog with the sent mail, we process only the addresses of recipients
            if (_app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).FolderPath == ((Outlook.Folder)mailItem.Parent).FolderPath)
            {
                if (_confFunc.IsEnableMarkedExternalRecipientsOnSentMail)
                {
                    IsExternalMail = IsExternalRecipients;
                }
                else
                {
                    IsExternalMail = (Outlook.MailItem o) => { return false; };
                }
            }
            else
            {
                IsExternalMail = IsExternalSender;
            }

            if (IsExternalMail(mailItem))
                MailMark(item);
        }

        /// <summary>
        /// Search for an external source address
        /// </summary>
        /// <returns></returns>
        public bool IsExternalSender(Outlook.MailItem mailItem)
        {
            if (mailItem == null || mailItem.Class != Outlook.OlObjectClass.olMail)
                return false;

            POCO.MailTransferPerson POCOSender = new POCO.MailTransferPerson()
            {
                Address = mailItem?.SenderEmailAddress?.ToUpper(),
                Name = mailItem?.SenderName?.ToUpper()
            };

            if (POCOSender.Address == null || _listHomeDomain.Count < 1)
                return false;

            if (!_listHomeDomain.Exists(m => POCOSender.Address.Contains(m.ToUpper())))
                return true;

            //Internal email
            return false;
        }

        /// <summary>
        /// Search for external destination recipients
        /// </summary>
        /// <returns></returns>
        public bool IsExternalRecipients(Outlook.MailItem mailItem)
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
                    //External email
                    return true;
                }
            }
            //Internal email
            return false;
        }

        /// <summary>
        /// Checking the email for compliance with the conditions of the specified configuration 
        /// for further processing by the add-on functions
        /// </summary>
        public override bool MailMatchesForProcessing(Outlook.MailItem mailItem)
        {
            //Compliance with the requirement to process only unread emails
            if (_confFunc.IsEnableAddititionalyHandlingOnlyUnread && !mailItem.UnRead)
                return false;

            return true;
        }

        /// <summary>
        /// The main function is action
        /// Marking the email
        /// </summary>
        private void MailMark(object item)
        {
            Outlook.MailItem mailItem = item as Outlook.MailItem;
            if (mailItem == null || mailItem.Class != Outlook.OlObjectClass.olMail)
                return;

            //If the email fits, then we perform the action
            if (!MailMatchesForProcessing(mailItem))
                return;

            try
            {
                if (!IsExistCategoryOnMailItem(mailItem, _confFunc.MailCategoryLabel))
                {
                    //Marking the email
                    if (String.IsNullOrEmpty(mailItem.Categories))
                        mailItem.Categories = _confFunc.MailCategoryLabel;
                    else
                        mailItem.Categories = $"{_confFunc.MailCategoryLabel}; {mailItem.Categories}";
                    mailItem.Save();
                }
            }
            catch { }
        }

        /// <summary>
        /// Does the email contain the specified marking category
        /// </summary>
        private bool IsExistCategoryOnMailItem(Outlook.MailItem mailItem, string categoryName)
        {
            string[] mailIsSetCategories = mailItem?.Categories?.Split(categorySeparator);

            //The email was received from outside
            if (mailIsSetCategories == null || mailIsSetCategories.Length == 0)
                return false;

            //The email was moved from another folder
            foreach (string elem in mailIsSetCategories)
            {
                if (String.Equals(elem.Trim(), categoryName))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Checking color categories in all user profile repositories
        /// </summary>
        private void VerifyCategoryOnOutlook(string categoryName, Outlook.OlCategoryColor categoryColor)
        {
            try
            {
                Outlook.Stores stores = _app.Session.Stores;
                Outlook.Categories currentCategories;
                bool isFound;
                foreach (Outlook.Store store in stores)
                {
                    currentCategories = store.Categories;
                    isFound = false;
                    foreach (Outlook.Category category in currentCategories)
                    {
                        if (
                            category.Name == categoryName
                            & category.Color == categoryColor
                            )
                        {
                            isFound = true;
                            break;
                        }
                    }
                    if (!isFound)
                        AddCategoryOnOutlook(currentCategories, categoryName, categoryColor);
                }
            }
            catch { }
        }

        /// <summary>
        /// Adding a category to the list of categories
        /// </summary>
        private void AddCategoryOnOutlook(Outlook.Categories categories, string categoryName, Outlook.OlCategoryColor categoryColor)
        {
            try
            {
                Outlook.Category category = categories[categoryName];
                if (category != null)
                {
                    category.Color = categoryColor;
                }
                else
                {
                    _ = categories.Add(
                    categoryName,
                    categoryColor,
                    Outlook.OlCategoryShortcutKey.olCategoryShortcutKeyNone);
                }
            }
            catch { }
        }
    }
}
