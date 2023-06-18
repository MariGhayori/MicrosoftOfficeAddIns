using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Diagnostics;

namespace WordOpenDocument
{
    public partial class ThisAddIn
    {
        private bool initialized = false;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.DocumentBeforeSave += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(this.Application_DocumentBeforeSave);
        }

        public void Application_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            Debug.Print("Document Before Save");

            IntPtr hWnd = Process.GetCurrentProcess().MainWindowHandle;

            Debug.Print("Window Handle: " + hWnd);
            Debug.Print("Document File Path: " + Doc.FullName);

            // In Case You want to stop a document from being Saved
            if (Doc.Name == "NewDoc.docx")
            {
                Cancel = true;
            }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
