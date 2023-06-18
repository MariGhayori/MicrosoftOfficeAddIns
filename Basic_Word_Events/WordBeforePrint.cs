using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Diagnostics;

namespace WordBeforePrint
{
    public partial class ThisAddIn
    {
        // Define a boolean "initialized" and set it to "false", it will be used later to stop the document from printing on right-click.
        private bool initialized = false;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // BeforePrint Event Occurs before the document is printed.
            Globals.ThisAddIn.Application.DocumentBeforePrint += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforePrintEventHandler(this.Application_DocumentBeforePrint);
        }

        public void Application_DocumentBeforePrint(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            // In This class, you can receive Document Data such as DocumentPath, Handle and etc
            Debug.Print("Document Before Print");

            IntPtr hWnd = Process.GetCurrentProcess().MainWindowHandle;

            Debug.Print("Window Handle: " + hWnd);
            Debug.Print("Document File Path: " + Doc.FullName);

            // In Case You want to stop a document from Printing
            if (Doc.Name == "NewDoc.docx")
            {
                Cancel = true;
            }
            // But this Won't Stop Printing if the user right-clicks on the document and chooses printing without opening the document.
            // For Your code to work on right-click too, follow the instructions
        }

        
        // Make The "InitializeCustomPrint" function and Call the BeforePrint Event in it.
        // Set the initialized to "true"
        private void InitializeCustomPrint()
        {
            initialized = true;
            Globals.ThisAddIn.Application.DocumentBeforePrint += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforePrintEventHandler(this.Application_DocumentBeforePrint);
        }

        // For this function to work You should go into "ThisAddIn.Designer.cs" of your project
        // Search for "Initialize()"
        // Then add the following code to the initialize():
        // this.InitializeCustomPrint();

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
