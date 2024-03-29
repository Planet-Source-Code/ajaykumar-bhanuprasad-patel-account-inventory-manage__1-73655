VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form show_sel_grp_smry 
   Caption         =   "Trial_balance_sheet"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox selected_from_list 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Text            =   "Select a group...,"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.ListBox combo_list 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   13800
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   119406593
      CurrentDate     =   40194
   End
   Begin MSFlexGridLib.MSFlexGrid grid_report 
      Height          =   6015
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   10610
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   480
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "GROUP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Select Group for Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6480
      TabIndex        =   1
      Top             =   360
      Width           =   6075
   End
End
Attribute VB_Name = "show_sel_grp_smry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub combo_list_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    selected_from_list.Text = combo_list.Text
    combo_list.Visible = False
    DTPicker1.Value = Date
    DTPicker1.SetFocus
    Label1.Caption = selected_group & "Closing Balance Summary as on " & DTPicker1.Value
    Call set_grid_report
End If
End Sub
Public Sub selected_group_summary()
Call make_trail_balance_summary
Call set_grid_report
End Sub
Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
selected_date = DTPicker1.Value
Call selected_group_summary
End Sub
Private Sub DTPicker1_Change()
selected_date = DTPicker1.Value
Call selected_group_summary
End Sub
Public Sub add_combo_list()
Call open_database
Call open_rs_lgr_prim_grp
Do Until rs_lgr_prim_grp.EOF
    combo_list.AddItem rs_lgr_prim_grp!lgr_prim_grp_name
    rs_lgr_prim_grp.MoveNext
Loop
Call open_rs_lgr_main_grp
Do Until rs_lgr_main_grp.EOF
    combo_list.AddItem rs_lgr_main_grp!lgr_main_grp_name
        If rs_lgr_main_grp!lgr_main_grp_alis <> "" Then
        combo_list.AddItem rs_lgr_main_grp!lgr_main_grp_alis
        End If
    rs_lgr_main_grp.MoveNext
Loop
End Sub
Private Sub Form_Load()
selected_date = Date
Label1.Caption = selected_procedure
DTPicker1.Value = Date
combo_list.Visible = False
Call add_combo_list
Call make_trail_balance_summary
'Call set_grid_report
End Sub
Public Sub set_grid_report()
Dim total_cr_balance
Dim total_dr_balance
total_cr_balance = 0
total_dr_balance = 0
    rep_ending_date = DTPicker1.Value
    grid_report.Clear
    grid_report.Rows = 1
    grid_report.Cols = 5
    grid_report.Font.Size = 12
    b = 0
    grid_report.Font.Size = 12
    grid_report.TextMatrix(b, 0) = ""
    grid_report.TextMatrix(b, 1) = "Ledger"
    grid_report.TextMatrix(b, 2) = "Group"
    grid_report.TextMatrix(b, 3) = "Dr.Amount"
    grid_report.TextMatrix(b, 4) = "Cr.Amount"
    grid_report.ColWidth(0) = 1 '800
    grid_report.ColWidth(1) = 6000
    grid_report.ColWidth(2) = 4000
    grid_report.ColWidth(3) = 3000
    grid_report.ColWidth(4) = 3000
    Dim x_grid_col
    Dim total_grid_width
    total_grid_width = 500
    For x_grid_col = 0 To grid_report.Cols - 1
        total_grid_width = total_grid_width + grid_report.ColWidth(x_grid_col)
    Next
grid_report.Width = total_grid_width
b = 1
Call open_rs_lgr_clsg_smr
rs_lgr_clsg_smr.Sort = "lgr_clsg_dtl_name"
Do Until rs_lgr_clsg_smr.EOF
If rs_lgr_clsg_smr!lgr_clsg_dtl_grup = selected_from_list.Text Or rs_lgr_clsg_smr!lgr_clsg_dtl_pgrp = selected_from_list.Text Then
                grid_report.Rows = grid_report.Rows + 1
                grid_report.TextMatrix(b, 0) = b
                grid_report.TextMatrix(b, 1) = rs_lgr_clsg_smr!lgr_clsg_dtl_name
                grid_report.TextMatrix(b, 2) = rs_lgr_clsg_smr!lgr_clsg_dtl_grup
                'rs_lgr_clsg_smr!lgr_clsg_dtl_pgrp = selected_primary_group
                'rs_lgr_clsg_smr!lgr_clsg_dtl_crpd = rs_lgr_main_dtl!lgr_main_dtl_crpd
                'rs_lgr_clsg_smr!lgr_clsg_dtl_cram = rs_lgr_main_dtl!lgr_main_dtl_cram
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal1 = rs_lgr_main_dtl!lgr_main_dtl_obl1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal2 = rs_lgr_main_dtl!lgr_main_dtl_obl2
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid1 = rs_lgr_main_dtl!lgr_main_dtl_osd1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid2 = rs_lgr_main_dtl!lgr_main_dtl_osd2
                'rs_lgr_clsg_smr!lgr_clsg_dtl_slun = rs_lgr_main_dtl!lgr_main_dtl_slun
                If rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "dr" Then
                grid_report.TextMatrix(b, 3) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal, "0.00")
                total_dr_balance = total_dr_balance + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal)
                ElseIf rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "cr" Then
                grid_report.TextMatrix(b, 4) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal, "0.00")
                total_cr_balance = total_cr_balance + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal)
                End If
    b = b + 1
End If
rs_lgr_clsg_smr.MoveNext
Loop
If b = 1 And total_dr_balance = 0 And total_cr_balance = 0 Then
MsgBox "There is no any account in this Group..,"
Exit Sub
End If
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 3) = "==================="
grid_report.TextMatrix(b, 4) = "==================="
b = b + 1
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 3) = Format(total_dr_balance, "0.00")
grid_report.TextMatrix(b, 4) = Format(total_cr_balance, "0.00")
b = b + 1
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 3) = "==================="
grid_report.TextMatrix(b, 4) = "==================="
End Sub
Private Sub grid_report_Click()
    selected_ledger = grid_report.TextMatrix(grid_report.Row, 1)
    ledger_clicked_from_other = 1
    selected_procedure = "show ledger account"
    shw_sel_lgr_dtl.Show
End Sub
Private Sub selected_from_list_Click()
    combo_list.Visible = True
    combo_list.Height = 2400
    combo_list.SetFocus
End Sub
