VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Calendar Two"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   3060
      Max             =   10
      TabIndex        =   7
      Top             =   360
      Value           =   1
      Width           =   1230
   End
   Begin VB.CheckBox chkSpan 
      Caption         =   "&S&pan empty calendar boxes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   6
      Top             =   2790
      Width           =   2850
   End
   Begin VB.TextBox txtBuild 
      Height          =   600
      Left            =   3150
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2925
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   555
      Left            =   3150
      TabIndex        =   4
      Top             =   2295
      Width           =   1100
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save HTML"
      Height          =   555
      Left            =   3150
      TabIndex        =   3
      Top             =   1350
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   2565
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboMonths 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   40
      Width           =   1725
   End
   Begin VB.ComboBox cboYears 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   45
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   45
      Width           =   1095
   End
   Begin VB.PictureBox picCal 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   135
      ScaleHeight     =   2130
      ScaleWidth      =   2805
      TabIndex        =   2
      Top             =   540
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Table Border:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3060
      TabIndex        =   9
      Top             =   45
      Width           =   1245
   End
   Begin VB.Label lblBorder 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3420
      TabIndex        =   8
      Top             =   720
      Width           =   405
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub cboMonths_Change()
    UpdateCalendar
End Sub

Private Sub cboMonths_Click()
    UpdateCalendar
End Sub

Private Sub cboYears_Change()
    UpdateCalendar
End Sub

Private Sub cboYears_Click()
    UpdateCalendar
End Sub

Private Sub chkSpan_Click()
    If chkSpan.Value <> 1 Then
        ShouldSpan = False
    Else
        ShouldSpan = True
    End If
End Sub

Private Sub Form_Load()
    Dim iCount As Integer, MaxWid As Long, MaxHig As Long
    
    LastPath = GetSetting("HTMLCalendar2", "Settings", "LastPath", App.Path)
    ShouldSpan = GetSetting("HTMLCalendar2", "Settings", "ShouldSpan", True)
    TableBorder = GetSetting("HTMLCalendar2", "Settings", "TableBorder", 1)

    HScroll1.Value = TableBorder
    If ShouldSpan = True Then chkSpan.Value = 1

    For iCount = 1990 To 2100
        cboYears.AddItem iCount
    Next

    For iCount = 1 To 12
        cboMonths.AddItem MonthName(iCount)
    Next

    cboYears.ListIndex = Year(Now) - 1990
    cboMonths.ListIndex = Month(Now) - 1
    
    MaxWid = picCal.TextWidth("--------------------")
    MaxHig = picCal.TextHeight(vbCrLf) * 8
    picCal.Width = MaxWid
    picCal.Height = MaxHig
    
    Load frmText

    UpdateCalendar

End Sub

Private Sub UpdateCalendar()

    Dim StartDay As Integer, iDay As Integer, iDate As Date
    Dim i As Integer, j As Integer
    Dim sRow As Integer
    
    '   Faster than erasing an array
    '
    For i = 1 To 6
        For j = 1 To 7
            CalGrid(i, j) = 0
        Next
    Next

    picCal.Cls
    iYear = Val(cboYears.List(cboYears.ListIndex))
    iMonth = cboMonths.ListIndex + 1
    StartDay = Weekday(DateSerial(iYear, iMonth, 1), vbSunday)

    picCal.Print " S  M  T  W  T  F  S"
    picCal.Print "--------------------"
    picCal.Print Space$((StartDay - 1) * 3); "01 ";
    If StartDay = 7 Then
        picCal.Print
        sRow = 2
    Else
        sRow = 1
    End If

    CalGrid(1, StartDay) = 1

    For iDay = 2 To 31
        iDate = DateSerial(iYear, iMonth, iDay)
        If Month(iDate) <> iMonth Then Exit For

        picCal.Print Format$(iDay, "00 ");

        CalGrid(sRow, Weekday(iDate)) = iDay

        If Weekday(iDate) = 7 Then
            picCal.Print
            sRow = sRow + 1
        End If
    Next
    picCal.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "HTMLCalendar2", "Settings", "LastPath", LastPath
    SaveSetting "HTMLCalendar2", "Settings", "ShouldSpan", ShouldSpan
    SaveSetting "HTMLCalendar2", "Settings", "TableBorder", TableBorder
    Unload frmText
End Sub

Private Sub btnSave_Click()

    CommonDialog1.CancelError = False
    On Error GoTo ErrHandler

    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt|HTM Files (*.htm)|*.htm;*.html"
    CommonDialog1.FilterIndex = 3
    CommonDialog1.InitDir = LastPath
    CommonDialog1.ShowSave
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    Dim i As Integer, j As Integer, ColSpan As Integer, MY As String
    Dim LastRowSum As Integer

    For i = Len(CommonDialog1.FileName) To 1 Step -1
        If Mid$(CommonDialog1.FileName, i, 1) = "\" Then
            LastPath = Left$(CommonDialog1.FileName, i)
            Exit For
        End If
    Next

    MY = cboMonths.List(cboMonths.ListIndex) & " " & Format$(iYear, "0000")
    txtBuild.Text = Replace(frmText.txtBlank.Text, "Month Year", MY, 1, 1, vbTextCompare)
    txtBuild.Text = Replace(txtBuild.Text, "border=2", "border=" & Format$(HScroll1.Value, "00"), 1, 1, vbTextCompare)
    
    If ShouldSpan = True Then

        For i = 1 To 7
            If CalGrid(1, i) <> 0 Then
                ColSpan = i - 1
                Exit For
            End If
        Next

        If ColSpan <> 0 Then
            txtBuild.Text = txtBuild.Text & "<td align=left valign=top colspan=" & Format$(ColSpan, "0") & ">&nbsp;</td>" & vbCrLf
        End If
    End If

    For i = 1 To 7
        If CalGrid(1, i) <> 0 Then
            txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
            txtBuild.Text = txtBuild.Text & Format$(CalGrid(1, i), "00") & vbCrLf
            txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
        Else
            If ShouldSpan = False Then
                txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
                txtBuild.Text = txtBuild.Text & "&nbsp;" & vbCrLf
                txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
            End If
        End If
    Next

    txtBuild.Text = txtBuild.Text & "</tr>" & vbCrLf & vbCrLf

    For i = 2 To 3
        txtBuild.Text = txtBuild.Text & "<tr>" & vbCrLf
        For j = 1 To 7
            txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
            txtBuild.Text = txtBuild.Text & Format$(CalGrid(i, j), "00") & vbCrLf
            txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
        Next

        txtBuild.Text = txtBuild.Text & "</tr>" & vbCrLf & vbCrLf

    Next

    For j = 1 To 7
        If CalGrid(4, j) = 0 Then
            If ShouldSpan = True Then
                ColSpan = 8 - j
                txtBuild.Text = txtBuild.Text & "<td align=left valign=top colspan=" & Format$(ColSpan, "0") & ">&nbsp;</td>" & vbCrLf
                Exit For
            Else
                txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
                txtBuild.Text = txtBuild.Text & "&nbsp;" & vbCrLf
                txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
            End If
        Else
            txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
            txtBuild.Text = txtBuild.Text & Format$(CalGrid(4, j), "00") & vbCrLf
            txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
        End If
    Next
    
    txtBuild.Text = txtBuild.Text & "</tr>" & vbCrLf & vbCrLf
        
    For j = 1 To 7
        If CalGrid(5, j) = 0 Then
            If ShouldSpan = True Then
                ColSpan = 8 - j
                txtBuild.Text = txtBuild.Text & "<td align=left valign=top colspan=" & Format$(ColSpan, "0") & ">&nbsp;</td>" & vbCrLf
                Exit For
            Else
                txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
                txtBuild.Text = txtBuild.Text & "&nbsp;" & vbCrLf
                txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
            End If
        Else
            txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
            txtBuild.Text = txtBuild.Text & Format$(CalGrid(5, j), "00") & vbCrLf
            txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
        End If
    Next

    txtBuild.Text = txtBuild.Text & "</tr>" & vbCrLf & vbCrLf
    
    For j = 1 To 7
        LastRowSum = LastRowSum + CalGrid(6, j)
    Next

    If LastRowSum <> 0 Then
        For j = 1 To 7
            If CalGrid(6, j) = 0 Then
                If ShouldSpan = True Then
                    ColSpan = 8 - j
                    txtBuild.Text = txtBuild.Text & "<td align=left valign=top colspan=" & Format$(ColSpan, "0") & ">&nbsp;</td>" & vbCrLf
                    Exit For
                Else
                txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
                txtBuild.Text = txtBuild.Text & "&nbsp;" & vbCrLf
                txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
                End If
            Else
                txtBuild.Text = txtBuild.Text & "<td align=left valign=top>" & vbCrLf
                txtBuild.Text = txtBuild.Text & Format$(CalGrid(6, j), "00") & vbCrLf
                txtBuild.Text = txtBuild.Text & "</td>" & vbCrLf
            End If
        Next
    End If

    txtBuild.Text = txtBuild.Text & "</table>" & vbCrLf & "</font>" & vbCrLf
    txtBuild.Text = txtBuild.Text & "</body>" & vbCrLf & "</html>" & vbCrLf
    
    If Dir$(CommonDialog1.FileName) <> "" Then Kill CommonDialog1.FileName
    Open CommonDialog1.FileName For Binary As #1
    Put #1, , txtBuild.Text
    Close

    Exit Sub
ErrHandler:
    Exit Sub

End Sub

Private Sub HScroll1_Change()
    lblBorder.Caption = Format$(HScroll1.Value, "00")
    TableBorder = HScroll1.Value
End Sub
