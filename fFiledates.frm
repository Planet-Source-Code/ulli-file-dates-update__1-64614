VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fFiledates 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "File Dates"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "fFiledates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   3390
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btOKRes 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   7155
      TabIndex        =   4
      ToolTipText     =   "Good Bye"
      Top             =   2670
      Width           =   1215
   End
   Begin VB.CommandButton btOKRes 
      Caption         =   "&Today"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   7155
      TabIndex        =   1
      ToolTipText     =   "Sets all calendars to today"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton btOKRes 
      Caption         =   "&Apply"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   7155
      TabIndex        =   3
      ToolTipText     =   "Apply dates to selected file"
      Top             =   2130
      Width           =   1215
   End
   Begin VB.CommandButton btOKRes 
      Cancel          =   -1  'True
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   7155
      TabIndex        =   2
      ToolTipText     =   "Reset"
      Top             =   1590
      Width           =   1215
   End
   Begin VB.CommandButton btBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8220
      TabIndex        =   0
      ToolTipText     =   "Browser"
      Top             =   435
      Width           =   345
   End
   Begin MSACAL.Calendar calFiledate 
      Height          =   2130
      Index           =   0
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1980
      _Version        =   524288
      _ExtentX        =   3493
      _ExtentY        =   3757
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2006
      Month           =   3
      Day             =   10
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   128
      FirstDay        =   2
      GridCellEffect  =   2
      GridFontColor   =   255
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   128
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar calFiledate 
      Height          =   2130
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1980
      _Version        =   524288
      _ExtentX        =   3493
      _ExtentY        =   3757
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2006
      Month           =   3
      Day             =   10
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   8388608
      FirstDay        =   2
      GridCellEffect  =   2
      GridFontColor   =   16711680
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar calFiledate 
      Height          =   2130
      Index           =   2
      Left            =   4800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1980
      _Version        =   524288
      _ExtentX        =   3493
      _ExtentY        =   3757
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2006
      Month           =   3
      Day             =   10
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   16384
      FirstDay        =   2
      GridCellEffect  =   2
      GridFontColor   =   32768
      GridLinesColor  =   0
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   49152
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   8145
      Top             =   555
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select File"
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Drop File or Folder here --->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Index           =   2
      Left            =   5760
      TabIndex        =   14
      Top             =   135
      Width           =   2370
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   7  'Diagonalkreuz
      Height          =   300
      Left            =   8220
      Top             =   90
      Width           =   345
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   150
      Index           =   1
      Left            =   1215
      TabIndex        =   13
      Top             =   3195
      Width           =   4590
   End
   Begin VB.Shape shpOK 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H00A0A0A0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   210
      Left            =   8445
      Shape           =   3  'Kreis
      Top             =   2250
      Width           =   210
   End
   Begin VB.Label lbCreated 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Created"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   825
      Width           =   1140
   End
   Begin VB.Label lbModified 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Last Modified"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2520
      TabIndex        =   11
      Top             =   825
      Width           =   1620
   End
   Begin VB.Label lbAccessed 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Last Accessed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   4800
      TabIndex        =   10
      Top             =   825
      Width           =   1725
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name and Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   135
      Width           =   1680
   End
   Begin VB.Label lbFileName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   435
      Width           =   7935
   End
End
Attribute VB_Name = "fFiledates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function OpenFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Private Enum FileConstants
    GENERIC_WRITE = &H40000000
    FILE_SHARE_READ = 1
    OPEN_EXISTING = 3
End Enum
#If False Then
Private GENERIC_READ, GENERIC_WRITE, FILE_SHARE_READ, OPEN_EXISTING
#End If

Private Enum TimeId
    Created = 0
    Modified = 1
    Accessed = 2
End Enum
#If False Then ':) Line inserted by Formatter
Private Created, Modified, Accessed ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

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

Private Const MAX_PATH              As Long = 260
Private Type WIN32_FIND_DATA
    dwFileAttributes                As Long
    ftCreationTime                  As FILETIME
    ftLastAccessTime                As FILETIME
    ftLastWriteTime                 As FILETIME
    nFileSizeHigh                   As Long
    nFileSizeLow                    As Long
    dwReserved0                     As Long
    dwReserved1                     As Long
    cFileName                       As String * MAX_PATH
    cAlternate                      As String * 14
End Type

Private hFile                       As Long
Private FileProps                   As WIN32_FIND_DATA
Private LocalDate                   As FILETIME
Private SysDate                     As SYSTEMTIME

Private SavedDate(Created To Accessed) As Date

Private Internal                    As Boolean
Private Const vbOrange              As Long = &HA0FF&
Private Const CNA                   As String = "Could not alter % date to "
Private Const DOR                   As String = "Date out of range"

Private Sub btBrowse_Click()

    shpOK.FillColor = &HA0A0A0
    cdl.ShowOpen
    lbFileName = cdl.FileName

End Sub

Private Sub btOKRes_Click(Index As Integer)

  Dim Signal As Long

    Select Case Index
      Case 0 'reset
        lbFileName = vbNullString
        Form_Load
        shpOK.FillColor = &HA0A0A0
        btOKRes(1).Enabled = True
      Case 1 'apply
        If GetLongPathName(lbFileName, 0, 0) Then
            Signal = vbGreen
            If PutFileDate(lbFileName, Created, SavedDate(Created)) = False Then
                Signal = vbOrange
                MsgBox Replace$(CNA, "%", "Created") & SavedDate(Created), , DOR
            End If
            If PutFileDate(lbFileName, Modified, SavedDate(Modified)) = False Then
                Signal = vbOrange
                MsgBox Replace$(CNA, "%", "Modified") & SavedDate(Modified), , DOR
            End If
            If PutFileDate(lbFileName, Accessed, SavedDate(Accessed)) = False Then
                Signal = vbOrange
                MsgBox Replace$(CNA, "%", "Accessed") & SavedDate(Accessed), , DOR
            End If
            shpOK.FillColor = Signal
            If Signal = vbOrange Then
                GetFileDates
            End If
          Else 'NOT GETLONGPATHNAME(LBFILENAME,...
            shpOK.FillColor = vbRed
        End If
      Case 2 'today
        Form_Load
      Case 3 'good bye
        Unload Me
    End Select

End Sub

Private Sub calFiledate_AfterUpdate(Index As Integer)

    Internal = False
    SavedDate(Index) = calFiledate(Index)

End Sub

Private Sub calFiledate_BeforeUpdate(Index As Integer, Cancel As Integer)

    Internal = True

End Sub

Private Sub calFiledate_NewMonth(Index As Integer)

  Dim tmp As Long

    If Not Internal Then
        tmp = Day(SavedDate(Index))
        calFiledate(Index).Day = 1 - (tmp = 1)
        calFiledate(Index).Day = tmp
        SavedDate(Index) = calFiledate(Index)
    End If

End Sub

Private Sub calFiledate_NewYear(Index As Integer)

  Dim tmp As Long

    If Not Internal Then
        tmp = Month(SavedDate(Index))
        calFiledate(Index).Month = 1 - (tmp = 1)
        calFiledate(Index).Month = tmp
        calFiledate_NewMonth Index
    End If

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

    SavedDate(Created) = Now
    SavedDate(Modified) = Now
    SavedDate(Accessed) = Now
    SavedDatesToCalendars
    lb(1) = "Do not use dates before " & Format$(DateSerial(1980, 1, 1), "Medium Date")
    If Len(Command$) > 2 Then
        lbFileName = Mid$(Command$, 2, Len(Command$) - 2)
        GetFileDates
    End If

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    shpOK.FillColor = &HA0A0A0
    On Error Resume Next
        lbFileName = Data.Files(1)
        If Err Then
            MsgBox "Files or Folders only please.", vbCritical, "Oops..."
        End If
    On Error GoTo 0

End Sub

Private Sub Form_Resize()

    If Len(Command$) > 2 Then
        lbFileName = Mid$(Command$, 2, Len(Command$) - 2)
    End If

End Sub

Private Function GetFileDate(FileName As String, ByVal Which As TimeId) As Date

  'returns a selected file date or Dec 31st, 9999 if file doesn't exist

    GetFileDate = DateSerial(9999, 12, 31)
    hFile = FindFirstFile(FileName, FileProps)
    If hFile <> -1 Then
        FindClose hFile
        With FileProps
            Select Case Which
              Case Created
                FileTimeToLocalFileTime .ftCreationTime, LocalDate
              Case Modified
                FileTimeToLocalFileTime .ftLastWriteTime, LocalDate
              Case Accessed
                FileTimeToLocalFileTime .ftLastAccessTime, LocalDate
              Case Else
                Exit Function '---> Bottom
            End Select
        End With 'FILEPROPS
        FileTimeToSystemTime LocalDate, SysDate
        With SysDate
            GetFileDate = DateSerial(.wYear, .wMonth, .wDay)
        End With 'SYSDATE
    End If

End Function

Private Sub GetFileDates()

    SavedDate(Created) = GetFileDate(lbFileName, Created)
    SavedDate(Modified) = GetFileDate(lbFileName, Modified)
    SavedDate(Accessed) = GetFileDate(lbFileName, Accessed)
    SavedDatesToCalendars

End Sub

Private Sub lbFileName_Change()

    If Len(lbFileName) Then
        GetFileDates
        btOKRes(1).Enabled = ((GetAttr(lbFileName) And vbDirectory) = 0)
        On Error Resume Next
            btOKRes(1).SetFocus
        On Error GoTo 0
    End If

End Sub

Private Function PutFileDate(FileName As String, ByVal Which As TimeId, NewDate As Date) As Boolean

  'modifies a selected file date, returns true for success

  Dim Attr As Long

    If IsDate(NewDate) Then
        hFile = FindFirstFile(FileName, FileProps)
        If hFile <> -1 Then
            FindClose hFile
            With SysDate
                .wYear = Year(NewDate)
                .wMonth = Month(NewDate)
                .wDay = Day(NewDate)
            End With 'SYSDATE
            SystemTimeToFileTime SysDate, LocalDate
            With FileProps
                Select Case Which
                  Case Created
                    LocalFileTimeToFileTime LocalDate, .ftCreationTime
                  Case Modified
                    LocalFileTimeToFileTime LocalDate, .ftLastWriteTime
                  Case Accessed
                    LocalFileTimeToFileTime LocalDate, .ftLastAccessTime
                  Case Else
                    Exit Function '---> Bottom
                End Select
                Attr = GetAttr(FileName)
                SetAttr FileName, Attr And Not vbReadOnly
                hFile = OpenFile(FileName, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, .dwFileAttributes, 0&)
                If hFile <> -1 Then
                    PutFileDate = SetFileTime(hFile, .ftCreationTime, .ftLastAccessTime, .ftLastWriteTime)
                    CloseHandle hFile
                End If
                SetAttr FileName, Attr
            End With 'FILEPROPS
        End If
    End If

End Function

Private Sub SavedDatesToCalendars()

    calFiledate(Created) = SavedDate(Created)
    calFiledate(Modified) = SavedDate(Modified)
    calFiledate(Accessed) = SavedDate(Accessed)

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Mrz-12 07:23)  Decl: 74  Code: 228  Total: 302 Lines
':) CommentOnly: 2 (0,7%)  Commented: 12 (4%)  Empty: 60 (19,9%)  Max Logic Depth: 5
