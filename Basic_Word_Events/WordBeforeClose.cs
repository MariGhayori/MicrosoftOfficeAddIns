using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Diagnostics;

namespace WordCloseDocument
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // DocumentOpen Event Occurs When a Document is Opened.
            Globals.ThisAddIn.Application.DocumentOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentOpenEventHandler(this.Application_DocumentBeforeClose);
        }

        // Add the DocumentOpen Function to "ThisAddIn" class so that when the event happens, this function will be fired.
        public void Application_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            // In This class, you can receive Document Data such as DocumentPath, Handle and etc
            Debug.Print("Document Closed");

            IntPtr hWnd = Process.GetCurrentProcess().MainWindowHandle;

            Debug.Print("Window Handle: " + hWnd);
            Debug.Print("Document File Path: " + Doc.FullName);

           // You can stop the document from closing using "Cancel = true;"
           if (Doc.Name = "sth.docx"){
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
