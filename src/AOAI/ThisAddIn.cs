using AOAI.Components;
using AOAI.Engine;

namespace AOAI
{
    public partial class ThisAddIn
    {
        AttentionSending _attentionSending;
        MarkingMail _markingMail;

        /// <summary>
        /// An analog of main. Processed when the add-on is loaded
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Config.LoadConfig();
            //Config.SaveConfig();

            //Notification when sending
            _attentionSending = new AttentionSending(application: this.Application);
            //Marking emails
            _markingMail = new MarkingMail(application: this.Application);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note. Outlook no longer issues this event. If there is a code that should be
            // executed when Outlook shuts down, see the article on the page https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Code created automatically by VSTO

        /// <summary>
        /// The required method to support the constructor — do not change 
        /// the contents of this method using the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
