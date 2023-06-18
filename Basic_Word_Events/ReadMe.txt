Basic c# VSTO for Word Event:
  - DocumentOpen
  - DocumentBeforePrint
  - DocumentBeforeSave
  - DocumentBeforeClose

  1. Document Open: This Event Occurs When a Document is Opened.
    - You have to call the event in "ThisAddIn_Startup"
    - Then You can Add the function to the "ThisAddIn" Class
    - In This class, you can receive Document Data such as DocumentPath, Handle and etc
    - You can also do some actions on the document like Making the file ReadOnly and etc

  2. Document Before Print:
    - You have to call the event in "ThisAddIn_Startup"
    - Then You can Add the function to the "ThisAddIn" Class 
    - In This class, you can receive Document Data such as DocumentPath, Handle and etc
    - In case you want to stop a document from printing, after your "Conditional Statements", you can put in below syntax: 
       Cancel = true;
    - Of course, this won't prevent printing if the user right-clicks on the document and chooses printing without opening the document.
    - For your code to work on right-click too, follow the given instructions

  3. Document Before Save:
    - You have to call the event in "ThisAddIn_Startup"
    - Then You can Add the function to the "ThisAddIn" Class 
    - In This class, you can receive Document Data such as DocumentPath, Handle and etc
    - In case you want to stop a document from saving, after your "Conditional Statements", you can put in below syntax: 
       Cancel = true;
