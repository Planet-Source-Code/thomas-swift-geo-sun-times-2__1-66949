Attribute VB_Name = "CSVFileReader"
Option Explicit
'Written by Thomas Swift - TAS Independent Programming May 26 2008
'I didn't see any reason to load the entire database into memory considering how large
'a database can get. So this module simply parses out the information and closes the
'file. Lots quicker, more accurate and keeps the memory free. Less code to. LOL
Private Type RECORDINFO
    City As String
    State As String
    Longitude As String
    Latitude As String
    TZ_Offset As String
    TZ_DST As Integer
End Type
Public Record_Info As RECORDINFO
Public Function GetZip(xPath As String, Zip As String) As Boolean
    Dim TheFile As Integer
    Dim OurBuffer As String
    On Error GoTo error
    GetZip = False
    If Dir(xPath) = vbNullString Then GoTo error
    TheFile = FreeFile()
    Open xPath For Input As #TheFile
    Do While Not EOF(TheFile)
        Input #TheFile, OurBuffer
        If Replace(Replace(OurBuffer, Chr(34), vbNullString), Chr(10), vbNullString) = Zip Then
            Input #TheFile, Record_Info.City
            Input #TheFile, Record_Info.State
            Input #TheFile, Record_Info.Latitude
            Input #TheFile, Record_Info.Longitude
            Input #TheFile, Record_Info.TZ_Offset
            Input #TheFile, Record_Info.TZ_DST
            GetZip = True
            Exit Do
        End If
        DoEvents
    Loop
    Close #TheFile
    Exit Function
    error:
    Close #TheFile
    MsgBox "File open error !"
End Function



