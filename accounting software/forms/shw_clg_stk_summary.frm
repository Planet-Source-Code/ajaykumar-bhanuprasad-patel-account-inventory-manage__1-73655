VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form shw_clg_stk_summary 
   Caption         =   "closing stock"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   10080
      Top             =   10440
   End
   Begin MSFlexGridLib.MSFlexGrid grid_stk_dtl 
      Height          =   6135
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   10821
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Label m_label 
      AutoSize        =   -1  'True
      Caption         =   "m_label"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   600
      TabIndex        =   2
      Top             =   7200
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Closing Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   2955
   End
End
Attribute VB_Name = "shw_clg_stk_summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    r_clr = 250
    g_clr = 50
    b_clr = 50

m_label.Caption = ""
Call open_database
Call open_rs_stk_item_lgr
Call set_stock_summary_grid
Call separation_of_all_inventory_to_inward_and_outward
Call search_closing_stock
Call enter_the_card_from_list
End Sub
Public Sub set_stock_summary_grid()

    grid_stk_dtl.RowHeightMin = 400
    grid_stk_dtl.Clear
    grid_stk_dtl.Rows = 2
    grid_stk_dtl.Cols = 7

    grid_stk_dtl.TextMatrix(0, 0) = "No"
    grid_stk_dtl.TextMatrix(0, 1) = "Item"
    grid_stk_dtl.TextMatrix(0, 2) = "Quantity"
    grid_stk_dtl.TextMatrix(0, 3) = "Rate"
    grid_stk_dtl.TextMatrix(0, 4) = "Amount"
    grid_stk_dtl.TextMatrix(0, 5) = "F.Val"
    grid_stk_dtl.TextMatrix(0, 6) = "Company"
    
    grid_stk_dtl.ColWidth(0) = 500
    grid_stk_dtl.ColWidth(1) = 4500
    grid_stk_dtl.ColWidth(2) = 1500
    grid_stk_dtl.ColWidth(3) = 1500
    grid_stk_dtl.ColWidth(4) = 1500
    grid_stk_dtl.ColWidth(5) = 1500
    grid_stk_dtl.ColWidth(6) = 2500

Dim temp_grid_col_no
Dim temp_grid_width
temp_grid_width = 0

For temp_grid_col_no = 0 To grid_stk_dtl.Cols - 1
temp_grid_width = temp_grid_width + grid_stk_dtl.ColWidth(temp_grid_col_no)
Next

grid_stk_dtl.Width = temp_grid_width + 800

grid_stk_dtl.Left = (Me.Width - grid_stk_dtl.Width) / 2
'grid_stk_dtl.Top = 2000
End Sub
Public Sub enter_the_card_from_list()
Call set_stock_summary_grid
    Dim rs_stk_clsg_srl_counter
    Dim grid_stk_row_no
    Dim total_inward
    Dim total_outward
    Dim temp_stock_balance
    grid_stk_row_no = 1
    grid_stk_dtl.Font.Size = 12

Call open_database
Call open_rs_stk_item_lgr
rs_stk_item_lgr.Sort = "stk_item_lgr_name"
Do Until rs_stk_item_lgr.EOF
    selected_stock_item_name = rs_stk_item_lgr!stk_item_lgr_name
    'Call open_database
    If rs_stk_clsg_srl.State = 1 Then rs_stk_clsg_srl.Close
    Call open_rs_stk_clsg_srl
    For rs_stk_clsg_srl_counter = 1 To rs_stk_clsg_srl.RecordCount
        If rs_stk_clsg_srl!stk_invt_clg_card = selected_stock_item_name Then
                temp_stock_balance = temp_stock_balance + (Val(rs_stk_clsg_srl!stk_invt_clg_edno) - Val(rs_stk_clsg_srl!stk_invt_clg_stno)) + 1
        End If
    rs_stk_clsg_srl.MoveNext
    Next
                    
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 0) = grid_stk_row_no
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = rs_stk_item_lgr!stk_item_lgr_name
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = temp_stock_balance
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = rs_stk_item_lgr!stk_item_lgr_rat1
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = Format(temp_stock_balance * Val(rs_stk_item_lgr!stk_item_lgr_rat1), "0.00")
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 5) = Format(rs_stk_item_lgr!stk_item_lgr_fcvl, "0.00")
    grid_stk_dtl.TextMatrix(grid_stk_row_no, 6) = rs_stk_item_lgr!stk_item_lgr_comp
    
    m_label.Caption = m_label.Caption & "  " & rs_stk_item_lgr!stk_item_lgr_name & "   (" & temp_stock_balance & ") " & Format(temp_stock_balance * Val(rs_stk_item_lgr!stk_item_lgr_rat1), "0.00") & "£      "
    temp_stock_balance = 0
    grid_stk_row_no = grid_stk_row_no + 1
    grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
    rs_stk_item_lgr.MoveNext
Loop

Dim all_item_total_stock_balance
Dim all_item_total_stock_balance_amount
all_item_total_stock_balance = 0
Dim grid_stk_dtl_counter
For grid_stk_dtl_counter = 1 To grid_stk_dtl.Rows - 1
    all_item_total_stock_balance = all_item_total_stock_balance + Val(grid_stk_dtl.TextMatrix(grid_stk_dtl_counter, 2))
    all_item_total_stock_balance_amount = all_item_total_stock_balance_amount + Val(grid_stk_dtl.TextMatrix(grid_stk_dtl_counter, 4))
Next
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "==========="
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = "==========="
        grid_stk_row_no = grid_stk_row_no + 1
        grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 1) = "Balance Quantity On" & Date & " "
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = all_item_total_stock_balance
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 3) = " Amount.."
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = Format(all_item_total_stock_balance_amount, "0.00")
        grid_stk_row_no = grid_stk_row_no + 1
        grid_stk_dtl.Rows = grid_stk_dtl.Rows + 1
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 2) = "==========="
        grid_stk_dtl.TextMatrix(grid_stk_row_no, 4) = "==========="
End Sub
Private Sub grid_stk_dtl_DblClick()
show_stock_item_by_click = 1
selected_stock_item_name = grid_stk_dtl.TextMatrix(grid_stk_dtl.Row, 1)
selected_procedure = "serial wise closing stock"
shw_item_wise_clg_stk.Show

End Sub

Private Sub Timer1_Timer()
If m_label.Left + m_label.Width <= 0 Then
    m_label.Left = Me.Width ' + m_label.Width
    If r_clr >= 250 Then
        g_clr = 50
        b_clr = b_clr + 50
    End If
    If g_clr >= 250 Then
        b_clr = 50
        r_clr = r_clr + 50
    End If
    If b_clr >= 250 Then
        r_clr = 50
        g_clr = g_clr + 50
    End If
End If
m_label.Left = m_label.Left - 250
m_label.ForeColor = RGB(r_clr, g_clr, b_clr)

End Sub
