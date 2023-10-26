using System;
using System.Windows.Forms;

namespace SandboxWordAddIn
{
    public partial class SandboxConversionAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            MessageBox.Show("Sandbox AddIn is active. Open any Word document to proceed.");
            Application.DocumentOpen += ApplicationOnDocumentOpen;
        }

        private void ApplicationOnDocumentOpen(Microsoft.Office.Interop.Word.Document interopDocument)
        {
            try
            {
                Microsoft.Office.Tools.Word.Document vstoObject = Globals.Factory.GetVstoObject(interopDocument);
                if (vstoObject == null)
                {
                    MessageBox.Show("Apologies, the current document instance could not be retrieved.");
                }
                else
                {
                    vstoObject.Select();
                    MessageBox.Show("Success: Document content selected successfully.");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Oops! Something went wrong: The application encountered an unexpected issue. " + e.Message);
                Console.WriteLine(e);
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}