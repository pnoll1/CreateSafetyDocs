Attribute VB_Name = "Module3"
Sub CreateSafetyDocs()
'Fills in safety docs & saves pdfs to job folder
'Dim project As String, customer As String, contact As String
project = ThisWorkbook.Worksheets("BASE").Cells(6, 3).Value
customer = ThisWorkbook.Worksheets("BASE").Cells(8, 3).Value
contact = ThisWorkbook.Worksheets("BASE").Cells(9, 3).Value
path = "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Safety"
pdfpath = "S:\Sicklesteel Cranes\Engineering\Clients\" & customer & "\" & project & "\5 Safety"

'Fill out FPP
Dim wrdapp As Word.Application
Set wrdapp = CreateObject("Word.Application")
wrdapp.Visible = True
fpp = wrdapp.Documents.Open(path & "\PTC - Fall Protection Plan.docx")
ActiveDocument.ContentControls(1).Range.Text = project
Set cc = ActiveDocument.SelectContentControlsByTitle("Customer")
ActiveDocument.ContentControls(2).Range.Text = customer 'Customer
Set cc = ActiveDocument.SelectContentControlsByTitle("Contact")
ActiveDocument.ContentControls(3).Range.Text = contact 'Contact
'export as pdf
ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfpath & "\PTC - Fall Protection Plan", ExportFormat:=wdExportFormatPDF
wrdapp.Quit (0)

'Fill out JHA

Workbooks.Open (path & "\Pacific Tower Crane - JHA.xlsx")
Worksheets(1).Cells(3, 1).Value = "Contract Title: " & project
Worksheets(1).Cells(3, 2).Value = "Contractor: " & customer
Worksheets(1).Cells(3, 3).Value = "Date: " & Date
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfpath & "\Pacific Tower Crane - JHA"
End Sub


