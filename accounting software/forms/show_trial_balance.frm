VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form show_trial_balance 
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   480
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
      Format          =   119209985
      CurrentDate     =   40194
   End
   Begin MSFlexGridLib.MSFlexGrid grid_report 
      Height          =   8295
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   14631
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorSel    =   -2147483648
      ForeColorSel    =   8388608
      SelectionMode   =   1
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
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Left            =   8760
      TabIndex        =   1
      Top             =   480
      Width           =   1515
   End
End
Attribute VB_Name = "show_trial_balance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
selected_date = DTPicker1.Value
Call make_trail_balance_summary
Call set_grid_report
End Sub

Private Sub DTPicker1_Change()
selected_date = DTPicker1.Value
Call make_trail_balance_summary
Call set_grid_report
End Sub

Private Sub Form_Load()
selected_date = Date
Label1.Caption = selected_procedure
DTPicker1.Value = Date
Call make_trail_balance_summary
Call set_grid_report
End Sub

Public Sub set_grid_report()
Dim total_cr_balance
Dim total_dr_balance
total_cr_balance = 0
total_dr_balance = 0

    rep_ending_date = DTPicker1.Value
    grid_report.Clear
    grid_report.Rows = 1
    grid_report.Cols = 4
    grid_report.Font.Size = 12
    
    b = 0
    
    grid_report.Font.Size = 12
    grid_report.TextMatrix(b, 0) = ""
    grid_report.TextMatrix(b, 1) = "Ledger"
    grid_report.TextMatrix(b, 2) = "Dr.Amount"
    grid_report.TextMatrix(b, 3) = "Cr.Amount"
    
    grid_report.ColWidth(0) = 800
    grid_report.ColWidth(1) = 10000
    grid_report.ColWidth(2) = 2000
    grid_report.ColWidth(3) = 2000
    
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
                
                'rs_lgr_clsg_smr!lgr_clsg_dtl_grup = rs_lgr_main_dtl!lgr_main_dtl_grup
                'rs_lgr_clsg_smr!lgr_clsg_dtl_pgrp = selected_primary_group
                'rs_lgr_clsg_smr!lgr_clsg_dtl_crpd = rs_lgr_main_dtl!lgr_main_dtl_crpd
                'rs_lgr_clsg_smr!lgr_clsg_dtl_cram = rs_lgr_main_dtl!lgr_main_dtl_cram
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal1 = rs_lgr_main_dtl!lgr_main_dtl_obl1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_bal2 = rs_lgr_main_dtl!lgr_main_dtl_obl2
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid1 = rs_lgr_main_dtl!lgr_main_dtl_osd1
                'rs_lgr_clsg_smr!lgr_clsg_dtl_sid2 = rs_lgr_main_dtl!lgr_main_dtl_osd2
                'rs_lgr_clsg_smr!lgr_clsg_dtl_slun = rs_lgr_main_dtl!lgr_main_dtl_slun
                
                If rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "dr" And rs_lgr_clsg_smr!lgr_clsg_dtl_tbal <> 0 Then
                
                grid_report.Rows = grid_report.Rows + 1
                grid_report.TextMatrix(b, 0) = b
                grid_report.TextMatrix(b, 1) = rs_lgr_clsg_smr!lgr_clsg_dtl_name
                
                grid_report.TextMatrix(b, 2) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal, "0.00")
                total_dr_balance = total_dr_balance + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal)
                b = b + 1

                ElseIf rs_lgr_clsg_smr!lgr_clsg_dtl_tsid = "cr" And rs_lgr_clsg_smr!lgr_clsg_dtl_tbal <> 0 Then
                
                grid_report.Rows = grid_report.Rows + 1
                grid_report.TextMatrix(b, 0) = b
                grid_report.TextMatrix(b, 1) = rs_lgr_clsg_smr!lgr_clsg_dtl_name
                
                
                grid_report.TextMatrix(b, 3) = Format(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal, "0.00")
                total_cr_balance = total_cr_balance + Val(rs_lgr_clsg_smr!lgr_clsg_dtl_tbal)
                b = b + 1

                End If
                
rs_lgr_clsg_smr.MoveNext
Loop
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = "==================="
grid_report.TextMatrix(b, 3) = "==================="
b = b + 1
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = Format(total_dr_balance, "0.00")
grid_report.TextMatrix(b, 3) = Format(total_cr_balance, "0.00")
b = b + 1
grid_report.Rows = grid_report.Rows + 1
grid_report.TextMatrix(b, 2) = "==================="
grid_report.TextMatrix(b, 3) = "==================="
End Sub

Private Sub grid_report_DblClick()
    selected_ledger = grid_report.TextMatrix(grid_report.Row, 1)
    ledger_clicked_from_other = 1
    selected_procedure = "show ledger account"
    shw_sel_lgr_dtl.Show
End Sub
