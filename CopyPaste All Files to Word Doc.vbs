Sub ProcessFiles()
    Dim Filename, Pathname As String
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Pathname = "C:\Users\nholt2\Desktop\Automation\Tables from R\Formatted\"
    Filename = Dir(Pathname & "*.xlsm")
    
    Do While Filename <> ""
        Set wb = Workbooks.Open(Pathname & Filename)
        wb.Activate
        DoWork wb
        
        wb.Close SaveChanges:=True
        Application.Wait (Now + TimeValue("00:00:01"))
        Filename = Dir()

    Loop
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub DoWork(wb As Workbook)
    With wb
        Application.Run ("FilenamePaste")
        Application.Run ("ExportMacro")
    End With
End Sub

Sub FilenamePaste()
    Dim obj As New DataObject
    Dim txt As String
    Dim WordApp As Word.Application
    Dim myDoc As Word.Document
    
    
    'Put some text inside a string variable
      txt = ActiveWorkbook.Name
    
    'Make object's text equal above string variable
      obj.SetText txt
    
    'Place DataObject's text into the Clipboard
      obj.PutInClipboard
      
    Set WordApp = GetObject(class:="Word.Application")
    WordApp.Visible = True
    Set myDoc = WordApp.Documents.Open("C:\Users\nholt2\Desktop\Automation\Macros\Survey Items.docx")
    WordApp.Activate
    SendKeys "^v"
    Application.Wait (Now + TimeValue("00:00:01"))
End Sub

Sub ExportMacro()
    Dim obj As New DataObject
    Dim txt As String
    Dim WordApp As Word.Application
    Dim myDoc As Word.Document
    
    'Copy the range Which you want to paste in a New Word Document
    Range("M1:S22").Copy
    
    Set WordApp = GetObject(class:="Word.Application")
    WordApp.Visible = True
    Set myDoc = WordApp.Documents.Open("C:\Users\nholt2\Desktop\Automation\Macros\Survey Items.docx")
    WordApp.Activate
    
    With WordApp
        .Selection.Paste
        .Visible = True
    End With
  
    
    With WordApp.Selection
        .InsertBreak Type:=7
    End With
     
End Sub