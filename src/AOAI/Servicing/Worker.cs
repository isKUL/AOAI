using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using AOAI.Components;

namespace AOAI.Servicing
{
    public static class Worker
    {
        /// <summary>
        /// Does the current user ignore the specified function of this add-on?
        /// </summary>
        public static bool IsIgnoredFunctionSuperUser(FunctionFeatures function)
        {
            string envUser = (String.IsNullOrEmpty(Environment.UserDomainName) ? Environment.UserName : $"{Environment.UserDomainName}\\{Environment.UserName}").ToUpper();
            Config conf = Config.Instance;
            POCO.SuperUser superUser = conf.ListIgnoredUsers.FirstOrDefault((user) => user.Name.ToUpper().Equals(envUser));
            if (superUser != null && superUser.ignoredFunctions != null && superUser.ignoredFunctions.Contains(function))
                return true;
            return false;
        }

        /// <summary>
        /// Checking for the presence of attachments in email
        /// </summary>
        public static bool IsWithAttachment(Outlook.MailItem mailItem)
        {
            if (mailItem == null || mailItem.Class != Outlook.OlObjectClass.olMail)
                return false;

            List<POCO.Attachment> POCOAttachmentList = new List<POCO.Attachment>();
            Outlook.Attachments attachments = mailItem.Attachments;

            foreach (Outlook.Attachment attachment in attachments)
            {
                //Игнорируем вложения внутри текста разметки HTML и RTF
                if (attachment.Size == 0)
                    continue;

                POCOAttachmentList.Add(
                    new POCO.Attachment()
                    {
                        FileName = attachment.FileName,
                        Size = attachment.Size
                    }
                );
            }
            return (POCOAttachmentList.Count > 0) ? true : false;
        }
    }
}
