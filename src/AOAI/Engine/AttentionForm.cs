using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using AOAI.Servicing;

namespace AOAI.Engine
{
    public partial class AttentionForm : Form
    {
        Outlook.MailItem _mailItem;
        List<bool> _wrapCancel;
        POCO.ConfigAttentionSending _confFunc;
        public AttentionForm(Outlook.MailItem mailItem, List<bool> wrapCancel, POCO.ConfigAttentionSending confFunc)
        {
            InitializeComponent();
            _mailItem = mailItem;
            _wrapCancel = wrapCancel;
            _confFunc = confFunc;
            LabelsLoad();
        }
        private void LabelsLoad()
        {
            this.Text = _confFunc?.FormTextLabel;
            this.btnSend.Text = _confFunc?.BtnSendLabel;
            this.btnCancel.Text = _confFunc?.BtnCancelLabel;
            LoadInformationText(_confFunc?.LabelInformation);
        }

        /// <summary>
        /// Text for the information window
        /// </summary>
        private void LoadInformationText(List<POCO.LabelTextObject> textInfo)
        {
            if (textInfo == null || textInfo.Count < 1)
                return;
            for (int i = 0; i < textInfo.Count; i++)
            {
                if (String.IsNullOrEmpty(textInfo[i].Text) || String.IsNullOrEmpty(textInfo[i].Color))
                    continue;
                try
                {
                    richTextBoxInformation.SelectionColor = Color.FromName(textInfo[i].Color);
                    richTextBoxInformation.AppendText(textInfo[i].Text);
                    if (i + 1 != textInfo.Count)
                        richTextBoxInformation.AppendText(Environment.NewLine);
                }
                catch { }
            }
        }

        /// <summary>
        /// Cancel sending
        /// </summary>
        private void Cancel_Click(object sender, EventArgs e)
        {
            _wrapCancel[0] = true;
            Button m = sender as Button;
            m?.FindForm()?.Close();
        }

        /// <summary>
        /// Confirm sending
        /// </summary>
        private void Send_Click(object sender, EventArgs e)
        {
            _wrapCancel[0] = false;
            Button m = sender as Button;
            m?.FindForm()?.Close();
        }
    }
}
