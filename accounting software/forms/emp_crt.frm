VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form empl_creat 
   ClientHeight    =   10950
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   14700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   14700
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   30
      Top             =   10680
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   19976
            MinWidth        =   7408
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "(C) Masino Sinaga (masino_sinaga@yahoo.com)"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "It's up to you..."
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "28/07/2010"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Date today"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1464
            MinWidth        =   1464
            TextSave        =   "22:15"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Time right now"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_name"
      Height          =   405
      Index           =   0
      Left            =   2760
      TabIndex        =   18
      Top             =   1200
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_rfnm"
      Height          =   405
      Index           =   1
      Left            =   2760
      TabIndex        =   19
      Top             =   1680
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_post"
      Height          =   405
      Index           =   2
      Left            =   2760
      TabIndex        =   20
      Top             =   2160
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_add1"
      Height          =   405
      Index           =   3
      Left            =   2760
      TabIndex        =   21
      Top             =   2640
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_add2"
      Height          =   405
      Index           =   4
      Left            =   2760
      TabIndex        =   22
      Top             =   3120
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_city"
      Height          =   405
      Index           =   5
      Left            =   2760
      TabIndex        =   23
      Top             =   3600
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_pncd"
      Height          =   405
      Index           =   6
      Left            =   2760
      TabIndex        =   24
      Top             =   4080
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_mobl"
      Height          =   405
      Index           =   7
      Left            =   2760
      TabIndex        =   25
      Top             =   4920
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_tel1"
      Height          =   405
      Index           =   8
      Left            =   2760
      TabIndex        =   26
      Top             =   5400
      Width           =   4230
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_ntin"
      Height          =   390
      Index           =   9
      Left            =   9000
      TabIndex        =   27
      Top             =   1185
      Width           =   2670
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_vist"
      Height          =   405
      Index           =   10
      Left            =   9000
      TabIndex        =   28
      Top             =   1680
      Width           =   2670
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_psno"
      Height          =   405
      Index           =   11
      Left            =   9000
      TabIndex        =   29
      Top             =   2160
      Width           =   2670
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_nino"
      Height          =   405
      Index           =   12
      Left            =   9000
      TabIndex        =   36
      Top             =   5520
      Width           =   2670
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_jodt"
      Height          =   405
      Index           =   13
      Left            =   9000
      TabIndex        =   31
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_hrrt"
      Height          =   405
      Index           =   14
      Left            =   9000
      TabIndex        =   32
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_wkrt"
      Height          =   405
      Index           =   15
      Left            =   9000
      TabIndex        =   33
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_sttm"
      Height          =   405
      Index           =   16
      Left            =   9000
      TabIndex        =   34
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "emp_main_dtl_entm"
      Height          =   405
      Index           =   17
      Left            =   9000
      TabIndex        =   35
      Top             =   5040
      Width           =   2655
   End
   Begin VB.PictureBox picButtons 
      Height          =   10305
      Left            =   12120
      ScaleHeight     =   10245
      ScaleWidth      =   2235
      TabIndex        =   42
      Top             =   600
      Width           =   2295
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   705
         Left            =   240
         TabIndex        =   37
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   705
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   705
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   705
         Left            =   240
         TabIndex        =   40
         Top             =   9360
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   2955
      Left            =   840
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   5212
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar prgBar1 
      Height          =   180
      Left            =   840
      TabIndex        =   45
      Top             =   9840
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   318
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox picStatBox 
      Height          =   600
      Left            =   840
      ScaleHeight     =   540
      ScaleWidth      =   10995
      TabIndex        =   46
      Top             =   10080
      Width           =   11055
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   350
         Left            =   120
         TabIndex        =   47
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Prev"
         Height          =   350
         Left            =   840
         TabIndex        =   48
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   350
         Left            =   9480
         TabIndex        =   49
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   350
         Left            =   10200
         TabIndex        =   50
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   51
         Top             =   120
         Width           =   7920
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Employee Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   52
      Top             =   240
      Width           =   10455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref Name:"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Post:"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Area:"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Postcode:"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile:"
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone:"
      Height          =   255
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality:"
      Height          =   255
      Index           =   9
      Left            =   7200
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Visa Status:"
      Height          =   255
      Index           =   10
      Left            =   7200
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Passport No:"
      Height          =   255
      Index           =   11
      Left            =   7200
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "NI No:"
      Height          =   255
      Index           =   12
      Left            =   7200
      TabIndex        =   12
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Joining Date:"
      Height          =   255
      Index           =   13
      Left            =   7200
      TabIndex        =   13
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Hourly Rate:"
      Height          =   255
      Index           =   14
      Left            =   7200
      TabIndex        =   14
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Weekly Rate:"
      Height          =   255
      Index           =   15
      Left            =   7200
      TabIndex        =   15
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Time:"
      Height          =   255
      Index           =   16
      Left            =   7200
      TabIndex        =   16
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending time:"
      Height          =   255
      Index           =   17
      Left            =   7200
      TabIndex        =   17
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblAngka 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5400
      TabIndex        =   43
      Top             =   8490
      Width           =   1455
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   44
      Top             =   8490
      Width           =   2655
   End
End
Attribute VB_Name = "empl_creat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private int_i As Integer

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim blnCancel As Boolean
Dim NumData As Integer
Dim intRecord As Integer
Dim intField As Integer

Private Sub Form_Load()
For int_i = 0 To 17
    lblLabels(i).FontSize = 15
    lblLabels(i).Height = 400
    txtFields(i).FontSize = 15
Next
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
On Error GoTo Message
Call open_database
Call open_rs_emp_main_dtl
Dim oText As TextBox
For Each oText In Me.txtFields
Set oText.DataSource = rs_emp_main_dtl
Next
Set grdDataGrid.DataSource = rs_emp_main_dtl.DataSource
mbDataChanged = False
LockTheForm
grdDataGrid.Enabled = True

    If rs_emp_main_dtl.RecordCount < 1 Then
    MsgBox "Recordset is empty. Please click Add button to add new record!", vbExclamation, "Empty Recordset"
    Exit Sub
    End If

LockTheForm
grdDataGrid.Enabled = True
grdDataGrid.TabStop = False
'SetButtons True
Exit Sub
Message:
MsgBox "hello"
MsgBox Err.Number & " - " & Err.Description
End
End Sub

Private Sub Message(strMessage As String)
  StatusBar1.Panels(1).Text = strMessage
End Sub
Private Sub cmdDataGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Adjust datagrid columns based on the longest field.")
End Sub


Private Sub cmdFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Go to the first record.")
End Sub

Private Sub cmdLast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Go to the last record.")
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Go to the next record.")
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Go to the previous record.")
End Sub
Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
  Response = -1
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If cmdUpdate.Enabled = True And cmdCancel.Enabled = True Then
     MsgBox "You have to save or cancel the changes " & vbCrLf & "that you have just made before quit!", vbExclamation, "Warning"
     cmdUpdate.SetFocus
     Cancel = -1
     Exit Sub
  End If

  If Not rs_emp_main_dtl Is Nothing Then _
    Set rs_emp_main_dtl = Nothing  'Clear memory from recordset
  
  If grdDataGrid.TabStop = True Then 'In order that prevent error from DataGrid...!
     txtFields(0).SetFocus
  End If
'  cnn.Close 'Close database
  Set cnn = Nothing  'Clear memory from database
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault 'Mouse pointer back to normal
End Sub

'Display the selected record in datagrid
Public Sub rs_emp_main_dtl_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  NumData = rs_emp_main_dtl.AbsolutePosition
  lblStatus.Caption = "Record number " & CStr(NumData) & " from " & rs_emp_main_dtl.RecordCount
  CheckNavigation
End Sub

Private Sub CheckNavigation()
  With rs_emp_main_dtl
   If (.RecordCount > 1) Then
      If (.BOF) Or (.AbsolutePosition = 1) Then
          cmdFirst.Enabled = False
          cmdPrevious.Enabled = False
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      ElseIf (.EOF) Or (.AbsolutePosition = .RecordCount) Then
          cmdNext.Enabled = False
          cmdLast.Enabled = False
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True

      Else
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      End If
   Else
      cmdFirst.Enabled = False
      cmdPrevious.Enabled = False
      cmdNext.Enabled = False
      cmdLast.Enabled = False
   End If
 End With
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With rs_emp_main_dtl
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    UnlockTheForm
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With
  grdDataGrid.Enabled = False  'In order that prevent error
  On Error Resume Next
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  LockTheForm
  grdDataGrid.Enabled = True
  If blnCancel = True Then
     Exit Sub
  End If
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rs_emp_main_dtl.CancelUpdate
  If mvBookMark > 0 Then
    rs_emp_main_dtl.Bookmark = mvBookMark
  Else
    If rs_emp_main_dtl.RecordCount > 0 Then rs_emp_main_dtl.MoveFirst
  End If
  LockTheForm    'Lock textbox
  grdDataGrid.Enabled = True
  mbDataChanged = False
End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Cancel the change or new record that have not been saved.")
End Sub
Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Save the change or new record.")
End Sub
Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  For int_i = 0 To 17
    If txtFields(i).Text = "" Then
       MsgBox "You have to fill in all textbox!", vbExclamation, "Validation"
       txtFields(i).SetFocus
       Exit Sub
     End If
  Next int_i
  'Update by using UpdateBatch. UpdateBatch will
  'automatically update all data in various fields type.
  rs_emp_main_dtl.UpdateBatch adAffectAll
  'Move pointer to last record if we just added data
  If mbAddNewFlag Then
    rs_emp_main_dtl.MoveLast
  End If
  If mbEditFlag Then
    rs_emp_main_dtl.MoveNext
    rs_emp_main_dtl.MovePrevious
  End If
  'Update all status
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  LockTheForm  'Lock textbox
  grdDataGrid.Enabled = True
  'Display the record position
  NumData = rs_emp_main_dtl.AbsolutePosition
  lblStatus.Caption = "Record number " & CStr(NumData) & " from " & rs_emp_main_dtl.RecordCount
  Exit Sub
UpdateErr:
  MsgBox Err.Number & " - " & Err.Description, vbCritical, "Error Occured"
End Sub
Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Quit from this program now.")
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
If adoFilter Is Nothing Then
   rs_emp_main_dtl.MoveFirst
Else
   adoFilter.MoveFirst
End If
  mbDataChanged = False
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub
Private Sub cmdLast_Click()
If adoFilter Is Nothing Then
   rs_emp_main_dtl.MoveLast
Else
   adoFilter.MoveLast
End If
  mbDataChanged = False
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub
Private Sub cmdNext_Click()
  On Error GoTo GoNextError
If adoFilter Is Nothing Then
   If Not rs_emp_main_dtl.EOF Then rs_emp_main_dtl.MoveNext
   If rs_emp_main_dtl.EOF And rs_emp_main_dtl.RecordCount > 0 Then
      Beep
      rs_emp_main_dtl.MoveLast
      MsgBox "This is the last record.", vbInformation, "Last Record"
   End If
Else
   If Not adoFilter.EOF Then adoFilter.MoveNext
   If adoFilter.EOF And adoFilter.RecordCount > 0 Then
      Beep
      adoFilter.MoveLast
      MsgBox "This is the last record.", vbInformation, "Last Record"
   End If
End If
  mbDataChanged = False
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError
If adoFilter Is Nothing Then
   If Not rs_emp_main_dtl.BOF Then rs_emp_main_dtl.MovePrevious
   If rs_emp_main_dtl.BOF And rs_emp_main_dtl.RecordCount > 0 Then
      Beep
      rs_emp_main_dtl.MoveFirst
      MsgBox "This is the first record.", _
             vbInformation, "First Record"
   End If
Else
   If Not adoFilter.BOF Then adoFilter.MovePrevious
   If adoFilter.BOF And adoFilter.RecordCount > 0 Then
      Beep
      adoFilter.MoveFirst
      MsgBox "This is the first record.", _
             vbInformation, "First Record"
   End If
End If
  mbDataChanged = False
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)

  cmdAdd.Enabled = bVal
  cmdUpdate.Enabled = Not bVal
  cmdCancel.Enabled = Not bVal
  cmdEdit.Enabled = bVal
  cmdDelete.Enabled = bVal
  cmdRefresh.Enabled = bVal
  cmdFind.Enabled = bVal
  cmdFilter.Enabled = bVal
  cmdSort.Enabled = bVal
  cmdBookmark.Enabled = bVal
  cmdDataGrid.Enabled = bVal
  cmdClose.Enabled = bVal

  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub picButtons_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index  'If we hit Enter, jump to next textbox
         Case 0 To 17
              If KeyAscii = 13 Then SendKeys "{Tab}"
  End Select
End Sub

'Lock textbox in order that we can't edit data
Private Sub LockTheForm()
  For int_i = 0 To 17
    txtFields(i).Locked = True
  Next int_i
  grdDataGrid.Enabled = False
End Sub

'Unlock textbox in order that we can edit data
Sub UnlockTheForm()

  For int_i = 0 To 17
    txtFields(i).Locked = False
  Next int_i
  grdDataGrid.Enabled = False
End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Find record (find first and find next).")
End Sub

Private Sub cmdFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Filter recordset.")
End Sub

Public Sub rsstrFindData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    NumData = rsstrFindData.AbsolutePosition
    lblStatus.Caption = "Record number " & CStr(NumData) & " from " & rsstrFindData.RecordCount
End Sub

Private Sub cmdSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Sort recordset.")
End Sub

Private Sub cmdBookmark_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call Message("Bookmark record so you can go back easily.")
End Sub

