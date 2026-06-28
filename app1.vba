Option Explicit
Dim RunTimer As Date
Sub Scheduler()
'
RunTimer = TimeValue("00:00:00")

Application.OnTime RunTimer, "Scheduler"
'MsgBox "Take a break.", vbInformation
Call Test

End Sub

'This is similar to app.vba. Other functionalities include the renaming of the months for conformity of the data.

Sub Conv_Mths()
Columns("G:G").Select

Selection.Replace What:="fÃ©vr", Replacement:="FEV", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

Selection.Replace What:="aoÃ»t", Replacement:="AUG", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

Selection.Replace What:="dÃ©c.", Replacement:="DEC", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

End Sub
Sub clear_range()

Range("A1:I1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("A1").Select

End Sub
Sub tfr_data(wksht As String)

Range("A1:I1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Windows("automated_tracker_Upstream.xlsx").Activate
Sheets(wksht).Select
Range("A1").Select
ActiveSheet.Paste

End Sub

Sub Test()

Dim olApp As Outlook.Application
Dim olNs As Namespace
Dim Fldr As MAPIFolder
Dim olMail As Variant
Dim olAtt As Outlook.Attachment
Dim mailitem As Outlook.mailitem
Dim strFname As String
Dim strTempFolder As String
strTempFolder = Environ("Temp") & Chr(92)

Set olApp = New Outlook.Application
Set olNs = olApp.GetNamespace("MAPI")
Set Fldr = olNs.GetDefaultFolder(olFolderInbox)

Dim rectime As Date
Dim theDate As Date
theDate = Now()

Dim i As Variant
Dim ti As String

'Dim sh As Shell32.Shell
'Dim DestinationFolder As Shell32.Folder
'Dim SourceFolder As Shell32.Folder

Dim zipFileName As String
Dim unzipFolderName As String
Dim objZipItems As FolderItems
Dim objZipItem As FolderItem

Dim wShApp As Shell
Set wShApp = CreateObject("Shell.Application")

Dim va As Variant
Dim va1 As Variant

Workbooks.OpenText Filename:="C:\up\automated_tracker_.xlsx"
ActiveWorkbook.Sheets("ARC_ETL").Activate
clear_range
            
ActiveWorkbook.Sheets("ARC_ETL_LOG").Activate
clear_range
            
ActiveWorkbook.Sheets("ARC_PAYS").Activate
clear_range
            
ActiveWorkbook.Sheets("ARC_TXT_LOG").Activate
clear_range
                  
For Each olMail In Fldr.Items
    If InStr(olMail.Subject, "some texts here related to the mail subject") <> 0 Then
        ti = olMail.ReceivedTime
        If Left(ti, 10) = Left(Now, 10) Then
            olMail.Display

            For Each olAtt In olMail.Attachments
                If Right(olAtt.Filename, 3) = "zip" Then
                    MsgBox "Le mail du jour est" & " : " & olAtt.Filename
                    olAtt.SaveAsFile "C:\up" & "\" & olAtt.Filename '
                    
                    zipFileName = "C:\up" & "\" & olAtt.Filename '
                    unzipFolderName = "C:\up" '
                    Set objZipItems = wShApp.Namespace(zipFileName).Items
                    
                    For Each objZipItem In wShApp.Namespace(zipFileName).Items
                        If InStr(objZipItem.Name, "ARC_ETL_Log.csv") <> 0 Then
                            wShApp.Namespace(unzipFolderName).CopyHere objZipItem
                        
                            Workbooks.OpenText Filename:="C:\up\" & objZipItem
                            
                            Columns("A:A").Select
                            'Application.CutCopyMode = False
                            
                            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
                            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
                            :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                            Array(7, 2), Array(8, 2), Array(9, 2)), TrailingMinusNumbers:=True
                            
                            Conv_Mths
                            tfr_data ("ARC_ETL_LOG")
                            
                        ElseIf InStr(objZipItem.Name, "ARC_TXT_Log.csv") <> 0 Then
                            wShApp.Namespace(unzipFolderName).CopyHere objZipItem
                        
                            Workbooks.OpenText Filename:="C:\up\" & objZipItem
                            
                            Columns("A:A").Select
                            'Application.CutCopyMode = False
                            
                            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
                            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
                            :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                            Array(7, 2), Array(8, 2), Array(9, 2)), TrailingMinusNumbers:=True
                            
                            Conv_Mths
                            tfr_data ("ARC_TXT_LOG")
                            
                        ElseIf InStr(objZipItem.Name, "_ARC_ETL.csv") <> 0 Then
                            wShApp.Namespace(unzipFolderName).CopyHere objZipItem
                        
                            Workbooks.OpenText Filename:="C:\up\" & objZipItem
                            
                            Columns("A:A").Select
                            'Application.CutCopyMode = False
                            
                            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
                            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
                            :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                            Array(7, 2), Array(8, 2), Array(9, 2)), TrailingMinusNumbers:=True
                            
                            Conv_Mths
                            tfr_data ("ARC_ETL")
                                
                        ElseIf InStr(objZipItem.Name, "_ARC_PAYS.csv") <> 0 Then ''''
                                                        wShApp.Namespace(unzipFolderName).CopyHere objZipItem
                        
                            Workbooks.OpenText Filename:="C:\up\" & objZipItem
                            
                            Columns("A:A").Select
                            
                            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
                            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
                            :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                            Array(7, 2), Array(8, 2), Array(9, 2)), TrailingMinusNumbers:=True
                            
                            Conv_Mths

                            tfr_data ("ARC_PAYS")

                            Workbooks("supervisionFR.xlsx").Worksheets("ARC_PAYS").Range("A1")

                            End If
                    Next
                End If
            Next olAtt
        End If
    End If
    
Next olMail

End Sub
