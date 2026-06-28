Option Explicit
Dim RunTimer As Date
Sub Scheduler()

RunTimer = TimeValue("00:00:00")

Application.OnTime RunTimer, "Scheduler"
Call Test

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

Dim fso As Object
Dim objFiles As Object
'Dim lngFileCount As Long

Dim MyPath As String
Dim MyPath1 As String
Dim MyPath2 As String
Dim MyFile As String
Dim MyFile1 As String
Dim MyFile2 As String
Dim LatestFile As String
Dim LatestFile1 As String
Dim LatestFile2 As String
Dim LatestDate As Date
Dim LatestDate1 As Date
Dim LatestDate2 As Date
Dim LMD As Date
Dim LMD1 As Date
Dim LMD2 As Date

Dim Path As String, Count As Integer, Filename As String

' This is a VBA script for an Excel file. When we click on it, it opens Outlook and search for a mail that contains some specific subject. It will download a zip file which contains Excel files.
' The script will open the necessary files and look for required data. Operations like filters, delimiters and cleaning are used in the script.
' The aim is to arrive at the desired data in each file and then transfer them automatically to another excel file in the local directory. This will update that excel file which already contains some formulae.
' The script helps to save time.
          
For Each olMail In Fldr.Items
    If InStr(olMail.Subject, "some texts here related to the mail subject") <> 0 Then
        ti = olMail.ReceivedTime
        If Left(ti, 10) = Left(Now, 10) Then
            olMail.Display
          
            Workbooks.OpenText Filename:="C:\Users\xxx\Desktop\Monitoring\supervisionRO.xlsx"
            ActiveWorkbook.Sheets("ARC_ETL").Activate
            Range("A1:I1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("A1").Select
            
            ActiveWorkbook.Sheets("ARC_PAYS").Activate ''''
            Range("A1:I1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("A1").Select
            '
            Workbooks.OpenText Filename:="C:\Users\xxx\Desktop\Monitoring\supervisionFR.xlsx"
            ActiveWorkbook.Sheets("ARC_ETL").Activate
            Range("A1:I1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("A1").Select
            
            ActiveWorkbook.Sheets("ARC_PAYS").Activate ''''
            Range("A1:I1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("A1").Select
            
            Workbooks.OpenText Filename:="C:\Users\xxx\Desktop\Monitoring\supervisionSK.xlsx"
            ActiveWorkbook.Sheets("ARC_ETL").Activate
            Range("A1:I1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("A1").Select
            
            ActiveWorkbook.Sheets("ARC_PAYS").Activate ''''
            Range("A1:I1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            Range("A1").Select
            
            For Each olAtt In olMail.Attachments
                If Right(olAtt.Filename, 3) = "zip" Then
                    MsgBox "Le mail du jour est" & " : " & olAtt.Filename

                    olAtt.SaveAsFile "C:\Users\xxx\Desktop\Monitoring" & "\" & olAtt.Filename '
                    
                    zipFileName = "C:\Users\xxx\Desktop\Monitoring" & "\" & olAtt.Filename '
                    unzipFolderName = "C:\Users\xxx\Desktop\Monitoring" '
                    Set objZipItems = wShApp.Namespace(zipFileName).Items
                    
                    '''
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    Set objFiles = fso.GetFolder(unzipFolderName).Files
                    '''
                    For Each objZipItem In wShApp.Namespace(zipFileName).Items
                        If InStr(objZipItem.Name, "ARC_") <> 0 Then
                            wShApp.Namespace(unzipFolderName).CopyHere objZipItem
                        
                            Workbooks.OpenText Filename:= _
                            "C:\Users\xxx\Desktop\Monitoring\" & objZipItem '
                            
                            Columns("A:A").Select
                            'Application.CutCopyMode = False
                            
                            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
                            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
                            :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                            Array(7, 2), Array(8, 2), Array(9, 2)), TrailingMinusNumbers:=True
                            
                            If InStr(objZipItem.Name, "_ARC_ETL.csv") <> 0 Then
                                va = Left(objZipItem, 16)
                                'MsgBox (va)
                                
                                ActiveWorkbook.Sheets(va).Activate
                                            Range("A1:I1").Select
                                            Range(Selection, Selection.End(xlDown)).Select                                
                                Selection.Copy _
                                Workbooks("supervisionRO.xlsx").Worksheets("ARC_ETL").Range("A1")
                                '''
                                Selection.Copy _
                                Workbooks("supervisionFR.xlsx").Worksheets("ARC_ETL").Range("A1")
                                
                                Selection.Copy _
                                Workbooks("supervisionSK.xlsx").Worksheets("ARC_ETL").Range("A1")
                                
                                ActiveWorkbook.Sheets(va).Activate
                                ActiveWorkbook.Close SaveChanges:=False
                                
                            ElseIf InStr(objZipItem.Name, "_ARC_PAYS.csv") <> 0 Then ''''
                                va1 = Left(objZipItem, 17)
                                ActiveWorkbook.Sheets(va1).Activate
                                            Range("A1:I1").Select
                                            Range(Selection, Selection.End(xlDown)).Select
                                Selection.Copy _
                                Workbooks("supervision Downstream RO.xlsx").Worksheets("ARC_PAYS").Range("A1")
                                '''
                                Selection.Copy _
                                Workbooks("supervision Downstream FR.xlsx").Worksheets("ARC_PAYS").Range("A1")
                                
                                Selection.Copy _
                                Workbooks("supervision Downstream SK.xlsx").Worksheets("ARC_PAYS").Range("A1")
                                
                                ActiveWorkbook.Sheets(va1).Activate
                                ActiveWorkbook.Close SaveChanges:=False
                                
                            End If
                        End If
                    Next
                    
                    For Each objZipItem In wShApp.Namespace(zipFileName).Items
                        If InStr(objZipItem.Name, "FR_DW_ARCA_CHK") <> 0 Then
                            wShApp.Namespace("C:\Users\xxx\Desktop\Monitoring\FR_DW_ARCA").CopyHere objZipItem
                        ElseIf InStr(objZipItem.Name, "RO_DW_ARCA_CHK") <> 0 Then
                            wShApp.Namespace("C:\Users\xxx\Desktop\Monitoring\RO_DW_ARCA").CopyHere objZipItem
                        ElseIf InStr(objZipItem.Name, "SK_DW_ARCA_CHK") <> 0 Then
                            wShApp.Namespace("C:\Users\xxx\Desktop\Monitoring\SK_DW_ARCA").CopyHere objZipItem
                        End If
                    Next
                    
                        MyPath = "C:\Users\xxx\Desktop\Monitoring\FR_DW_ARCA\"
                        MyPath1 = "C:\Users\xxx\Desktop\Monitoring\RO_DW_ARCA\"
                        MyPath2 = "C:\Users\xxx\Desktop\Monitoring\SK_DW_ARCA\"
                        
                            MyFile = Dir(MyPath & "*.csv", vbNormal)
                            
                            Path = MyPath & "\*"
                            Filename = Dir(Path)
                            Count = 1
                            Do While MyFile <> ""
                                
                                Filename = Dir()
                                'MsgBox Filename
                                Count = Count + 1                                
                                If Count = 3 Then
                                    Workbooks.Open MyPath & Filename
                                    Columns("A:A").Select
                                    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                                    Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                                    :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                                    Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2), Array(11, 2)), _
                                    TrailingMinusNumbers:=True
                                    Exit Do
                                End If
                            Loop
                        MyFile = Dir(MyPath1 & "*.csv", vbNormal)
                        Path = MyPath1 & "\*"
                        
                        Filename = Dir(Path)
                        Count = 1
                        Do While MyFile <> ""
                            Filename = Dir()
                            Count = Count + 1
                            If Count = 3 Then
                                Workbooks.Open MyPath1 & Filename
1                               Columns("A:A").Select
                                Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                                Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                                :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                                Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2), Array(11, 2)), _
                                TrailingMinusNumbers:=True
                                Exit Do
                            End If
                        Loop
                        MyFile = Dir(MyPath2 & "*.csv", vbNormal)
                        Path = MyPath2 & "\*"
                        
                        Filename = Dir(Path)
                        Count = 1
                        Do While MyFile <> ""
                            Filename = Dir()
                            Count = Count + 1
                            If Count = 3 Then
                                Workbooks.Open MyPath2 & Filename
                                Columns("A:A").Select
                                Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                                Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                                :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
                                Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2), Array(11, 2)), _
                                TrailingMinusNumbers:=True
                                Exit Do
                            End If
                        Loop
                    End If
            Next olAtt
        End If
    End If

Next olMail

End Sub
