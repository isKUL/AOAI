using Outlook = Microsoft.Office.Interop.Outlook;
using AOAI.Servicing;

namespace AOAI.Engine
{
    /// <summary>
    /// Abstract class for functionality classes
    /// </summary>
    abstract class FunctionAdd_Ins
    {
        abstract public void LoadConfig();
        abstract public FunctionFeatures FunctionFeature { get; }
        abstract public bool MailMatchesForProcessing(Outlook.MailItem mailItem);
    }
}
