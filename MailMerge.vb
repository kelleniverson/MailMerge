Option Explicit

Const FOLDER_SAVED As String = "J:\Compensation\Sales Comp Documents\2020\RD\Comp Docs- October\October 2020 RD Comp Docs-UNSIGNED\"

Const SOURCE_FILE_PATH As String = "J:\Commissions\2020\Sales\Adhoc & Analyses\October Commission Plan Setup-MM.xlsx"

Sub MailMerge_Automation()
Dim MainDoc As Document, TargetDoc As Document
Dim recordNumber As Long, totalRecord As Long

Set MainDoc = ThisDocument
With MainDoc.MailMerge

    .OpenDataSource Name:=SOURCE_FILE_PATH, SQLStatement:="SELECT * FROM[October Summary$]"
    
    totalRecord = .DataSource.RecordCount
    
    For recordNumber = 1 To totalRecord
     With .DataSource
        .ActiveRecord = recordNumber
        .FirstRecord = recordNumber
        .LastRecord = recordNumber
    End With
    
    .Destination = wdSendToNewDocument
    .Execute False
    
    Set TargetDoc = ActiveDocument
       TargetDoc.SaveAs2 FOLDER_SAVED & .DataSource.DataFields("RD_Name").Value & ".docx", wdFormatDocumentDefault
       
       TargetDoc.Close False
       
       
    Set TargetDoc = Nothing
    
       
Next recordNumber

End With

Set MainDoc = Nothing

End Sub