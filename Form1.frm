VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Geo Sun Times"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   -15
      ScaleHeight     =   405
      ScaleWidth      =   4830
      TabIndex        =   7
      Top             =   0
      Width           =   4860
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1335
         TabIndex        =   9
         Text            =   "85219"
         Top             =   38
         Width           =   1650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         Height          =   270
         Left            =   3030
         TabIndex        =   8
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sun"
      Height          =   1695
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   4635
      Begin VB.Label Label4 
         Caption         =   "Sunrise:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "Noon:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Sunset"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1380
         Width           =   555
      End
      Begin VB.Label lblSunrise 
         BackColor       =   &H80000018&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   3555
      End
      Begin VB.Label lblSunNoon 
         BackColor       =   &H80000018&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   3555
      End
      Begin VB.Label lblSunset 
         BackColor       =   &H80000018&
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   1320
         Width           =   3555
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(32) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(32) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
Private MyCap As String
Private cSun As clsSunrise
Private Sub GetGeoData()
lblSunrise.Caption = ""
lblSunNoon.Caption = ""
lblSunset.Caption = ""
    If IsNumeric(Text1.Text) = False Or Len(Text1.Text) <> 5 Then
        MsgBox "Zip code is not formated properly !"
        Exit Sub
    End If
    If VerifyGeoData = False Then
        If SetGeoData() = True Then
            GetGeoData
        Else
            MsgBox "Unable to set Geo Location, check your internet connection and try again !"
        End If
    Else
        Me.Caption = MyCap & " - " & GetSetting("GEO", "Location", "City", vbNullString)
        Set cSun = New clsSunrise
                
        cSun.Latitude = CDbl(GetSetting("GEO", "Location", "Latitude", ""))
        cSun.Longitude = CDbl(GetSetting("GEO", "Location", "Longitude", ""))
        cSun.TimeZone = CDbl(GetTimeZoneOffset)
        
        cSun.DateDay = Now
        cSun.DaySavings = 0
        cSun.CalculateSun
        
        Frame1.Caption = GetSetting("GEO", "Location", "City", "") & " " & GetSetting("GEO", "Location", "State", "")
        lblSunrise.Caption = cSun.Sunrise
        lblSunNoon.Caption = cSun.SolarNoon 'SunTransit
        lblSunset.Caption = cSun.Sunset
        Set cSun = Nothing
    End If
End Sub
Private Function VerifyGeoData() As Boolean
    VerifyGeoData = True
    If GetSetting("GEO", "Location", "Latitude", "") = "" Or _
            GetSetting("GEO", "Location", "Longitude", "") = "" Or _
            GetSetting("GEO", "Location", "City", "") = "" Or _
            GetSetting("GEO", "Location", "State", "") = "" Or _
            GetSetting("GEO", "Location", "Zip", "") <> Text1.Text _
            Then VerifyGeoData = False
End Function
Private Function SetGeoData() As Boolean
    Dim MyDoc As String
    SetGeoData = False
    Me.Caption = MyCap & " - Searching"
    If GetZip(App.Path & "\zipcode.csv", Format(Text1.Text, "00000")) = False Then
        MsgBox "Address dont exsist !"
    Else
    
    SaveSetting "GEO", "Location", "Latitude", Record_Info.Latitude
    SaveSetting "GEO", "Location", "Longitude", Record_Info.Longitude
    SaveSetting "GEO", "Location", "City", Record_Info.City
    SaveSetting "GEO", "Location", "State", Record_Info.State
    SaveSetting "GEO", "Location", "Zip", Format(Text1.Text, "00000")
    'SaveSetting "GEO", "Location", "Country", Record_Info.State
    SetGeoData = True
    End If
    Me.Caption = MyCap
End Function
Private Sub Command1_Click()
    Picture1.SetFocus
    GetGeoData
End Sub
Private Function GetTimeZoneOffset() As Integer
    Dim tz As TIME_ZONE_INFORMATION
    Dim lRV As Long
    Dim iRV As Integer
    lRV = GetTimeZoneInformation(tz)
    ' offset is indicated in minutes
    iRV = tz.Bias
    GetTimeZoneOffset = iRV / -60
End Function
Private Sub Form_Load()
    MyCap = Me.Caption
    Text1.Text = GetSetting("GEO", "Location", "Zip", "Enter Zip Here")
    GetGeoData
End Sub
