VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form creat_stock_lgr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Detail.....,"
      Height          =   375
      Left            =   7920
      TabIndex        =   24
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Detail.....,"
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   400
      Left            =   1560
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   6600
      Width           =   1200
   End
   Begin VB.TextBox Text10 
      Height          =   400
      Left            =   3960
      TabIndex        =   18
      Text            =   "Text10"
      Top             =   6600
      Width           =   1200
   End
   Begin VB.TextBox Text12 
      Height          =   400
      Left            =   6600
      TabIndex        =   19
      Text            =   "Text12"
      Top             =   6600
      Width           =   1200
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7560
      TabIndex        =   14
      Text            =   "Cr/Dr"
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   7560
      TabIndex        =   45
      Text            =   "Cr/Dr"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   400
      Left            =   7080
      TabIndex        =   43
      Text            =   "Text18"
      Top             =   9960
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Height          =   400
      Left            =   3120
      TabIndex        =   42
      Text            =   "Text17"
      Top             =   9960
      Width           =   2715
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   3120
      TabIndex        =   8
      Text            =   "Select a ledger to edit"
      Top             =   2280
      Width           =   5535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7560
      TabIndex        =   13
      Text            =   "Select_Group"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   400
      Left            =   7080
      TabIndex        =   39
      Text            =   "Text16"
      Top             =   9480
      Width           =   1545
   End
   Begin VB.TextBox Text15 
      Height          =   400
      Left            =   3120
      TabIndex        =   38
      Text            =   "Text15"
      Top             =   9480
      Width           =   1200
   End
   Begin VB.TextBox Text14 
      Height          =   400
      Left            =   3120
      TabIndex        =   37
      Text            =   "Text14"
      Top             =   9000
      Width           =   1200
   End
   Begin VB.TextBox Text13 
      Height          =   400
      Left            =   6600
      TabIndex        =   23
      Text            =   "Text13"
      Top             =   7440
      Width           =   1200
   End
   Begin VB.TextBox Text11 
      Height          =   400
      Left            =   3960
      TabIndex        =   22
      Text            =   "Text11"
      Top             =   7440
      Width           =   1200
   End
   Begin VB.TextBox Text9 
      Height          =   400
      Left            =   1560
      TabIndex        =   21
      Text            =   "Text9"
      Top             =   7440
      Width           =   1200
   End
   Begin VB.TextBox Text7 
      Height          =   400
      Left            =   3120
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   5640
      Width           =   5500
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6240
      TabIndex        =   27
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   26
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   495
      Left            =   2400
      TabIndex        =   25
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   400
      Left            =   3120
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   5160
      Width           =   5500
   End
   Begin VB.TextBox Text5 
      Height          =   400
      Left            =   3120
      TabIndex        =   28
      Text            =   "Text5"
      Top             =   4680
      Width           =   5500
   End
   Begin VB.TextBox Text4 
      Height          =   400
      Left            =   3120
      TabIndex        =   29
      Text            =   "Text4"
      Top             =   4200
      Width           =   5500
   End
   Begin VB.TextBox Text3 
      Height          =   400
      Left            =   3120
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   3720
      Width           =   5500
   End
   Begin VB.TextBox Text2 
      Height          =   400
      Left            =   3120
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   3240
      Width           =   5500
   End
   Begin VB.TextBox Text1 
      Height          =   400
      Left            =   3120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2760
      Width           =   5500
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8895
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
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
      Left            =   360
      TabIndex        =   49
      Top             =   6600
      Width           =   1200
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
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
      Left            =   3000
      TabIndex        =   48
      Top             =   6600
      Width           =   1200
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   47
      Top             =   6600
      Width           =   1200
   End
   Begin VB.Label lbl_Heading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Item Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   1680
      Width           =   8415
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
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
      Left            =   6000
      TabIndex        =   44
      Top             =   9960
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
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
      Left            =   360
      TabIndex        =   41
      Top             =   9960
      Width           =   2775
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
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
      Left            =   360
      TabIndex        =   40
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   36
      Top             =   9480
      Width           =   960
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   35
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
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
      Left            =   3000
      TabIndex        =   34
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
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
      Left            =   360
      TabIndex        =   33
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
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
      Left            =   360
      TabIndex        =   32
      Top             =   7080
      Width           =   3000
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   6240
      Width           =   3000
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
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
      Left            =   360
      TabIndex        =   30
      Top             =   5760
      Width           =   3000
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   3000
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   6
      Top             =   5280
      Width           =   3000
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   5
      Top             =   4800
      Width           =   3000
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   3000
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lbl_add 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   8415
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   7335
   End
End
Attribute VB_Name = "creat_stock_lgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_exit_Click()
Unload Me
End Sub
Private Sub cmd_save_Click()
'selected_procedure = "Stock_item_ledger_edit"
'selected_procedure = "Stock_item_ledger_creat"
Dim selected_stock_item_alias
If Text2.Text = "" Or Text2.Text = " " Then
    selected_stock_item_alias = "XXXXXXXXX"
Else
    selected_stock_item_alias = Text2.Text
End If
'check the values
If Text1.Text = "" Or Combo1.Text = "" Or Combo1.Text = "Select Group" Then
    MsgBox "You have not entered any value...!!!"
    Exit Sub
End If
'check for duplicate
If selected_procedure = "Stock_item_ledger_edit" Then
            Dim named_ledgers
            named_ledgers = 0
            Call open_database
            Call open_rs_stk_item_lgr
            
            Do Until rs_stk_item_lgr.EOF
            
                If rs_stk_item_lgr!stk_item_lgr_name = Text1.Text Or rs_stk_item_lgr!stk_item_lgr_alis = Text1.Text Or _
                    rs_stk_item_lgr!stk_item_lgr_name = selected_stock_item_alias Or rs_stk_item_lgr!stk_item_lgr_alis = selected_stock_item_alias Then
                        named_ledgers = named_ledgers + 1
                End If
            rs_stk_item_lgr.MoveNext
            Loop
            
            If named_ledgers > 1 And selected_stock_item_alias = "XXXXXXXXX" Then
                                MsgBox "This Ledger is already exist...!!!", vbOKOnly, "Duplicate"
                                Call arrange_form_item
                                Exit Sub
            ElseIf named_ledgers > 2 And selected_stock_item_alias <> "XXXXXXXXX" Then
                                MsgBox "This Ledger is already exist...!!!", vbOKOnly, "Duplicate"
                                Call arrange_form_item
                                Exit Sub
            End If

            'save
            Call open_database
            Call open_rs_stk_item_lgr
            Do Until rs_stk_item_lgr.EOF
            If rs_stk_item_lgr!stk_item_lgr_name = Combo3.Text Or rs_stk_item_lgr!stk_item_lgr_alis = Combo3.Text Then
                    
                    rs_stk_item_lgr!stk_item_lgr_name = Text1.Text
                    rs_stk_item_lgr!stk_item_lgr_alis = Text2.Text
                    rs_stk_item_lgr!stk_item_lgr_comp = Text3.Text
                    'rs_stk_item_lgr!stk_item_lgr_grup = Text4.Text
                    rs_stk_item_lgr!stk_item_lgr_unit = Combo2.Text
                    
                    rs_stk_item_lgr!stk_item_lgr_fcvl = Text6.Text
                    rs_stk_item_lgr!stk_item_lgr_disc = Text7.Text
                        
                        If Text8.Text = "" Then Text8.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_qnt1 = Text8.Text
                        If Text9.Text = "" Then Text9.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_qnt2 = Text9.Text
                        
                        If Text10.Text = "" Then Text10.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_rat1 = Text10.Text
                        If Text11.Text = "" Then Text11.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_rat2 = Text11.Text
                        
                        If Text12.Text = "" Then Text12.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_amt1 = Text12.Text
                        If Text13.Text = "" Then Text13.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_amt2 = Text13.Text
                        
                        rs_stk_item_lgr!stk_item_lgr_grup = Combo1.Text
                    
                    rs_stk_item_lgr.UpdateBatch
                End If
            rs_stk_item_lgr.MoveNext
            Loop
ElseIf selected_procedure = "Stock_item_ledger_creat" Then
                Call open_database
                Call open_rs_stk_item_lgr
                
                Do Until rs_stk_item_lgr.EOF
                    If rs_stk_item_lgr!stk_item_lgr_name = Text1.Text Or rs_stk_item_lgr!stk_item_lgr_alis = Text1.Text Or _
                    rs_stk_item_lgr!stk_item_lgr_name = selected_stock_item_alias Or rs_stk_item_lgr!stk_item_lgr_alis = selected_stock_item_alias Then
                        MsgBox "This Ledger is already exist...!!!", vbOKOnly, "Duplicate"
                        Call arrange_form_item
                        Exit Sub
                    End If
                rs_stk_item_lgr.MoveNext
                Loop
                
                'save
                Call open_database
                Call open_rs_stk_item_lgr
                        rs_stk_item_lgr.AddNew
                        
                    rs_stk_item_lgr!stk_item_lgr_name = Text1.Text
                    rs_stk_item_lgr!stk_item_lgr_alis = Text2.Text
                    rs_stk_item_lgr!stk_item_lgr_comp = Text3.Text
                    'rs_stk_item_lgr!stk_item_lgr_grup = Text4.Text
                    rs_stk_item_lgr!stk_item_lgr_unit = Combo2.Text
                    'Combo2.Text = rs_stk_item_lgr!stk_item_lgr_unit
                    rs_stk_item_lgr!stk_item_lgr_fcvl = Text6.Text
                    rs_stk_item_lgr!stk_item_lgr_disc = Text7.Text
                        
                        If Text8.Text = "" Then Text8.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_qnt1 = Text8.Text
                        If Text9.Text = "" Then Text9.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_qnt2 = Text9.Text
                        
                        If Text10.Text = "" Then Text10.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_rat1 = Text10.Text
                        If Text11.Text = "" Then Text11.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_rat2 = Text11.Text
                        
                        If Text12.Text = "" Then Text12.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_amt1 = Text12.Text
                        If Text13.Text = "" Then Text13.Text = 0
                        rs_stk_item_lgr!stk_item_lgr_amt2 = Text13.Text
                        
                        rs_stk_item_lgr!stk_item_lgr_grup = Combo1.Text
                        
                        rs_stk_item_lgr.UpdateBatch
End If
Call arrange_form_item
End Sub
Private Sub Combo1_Change()
'Combo1.Text = ""
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
Call open_database
Call open_rs_stk_item_grp
    Do Until rs_stk_item_grp.EOF
        If rs_stk_item_grp!stk_item_grp_alis = Combo1.Text Then
        selected_stock_group = rs_stk_item_grp!stk_item_grp_name
        Combo1.Text = selected_stock_group
        'MsgBox rs_stk_item_grp!stk_item_grp_name
        Exit Sub
        End If
    rs_stk_item_grp.MoveNext
    Loop
End Sub

Private Sub Combo2_Change()
'Combo2.Text = ""
End Sub
Private Sub Combo3_Change()
'Combo3.Text = ""
End Sub
Private Sub Combo3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
'Text14.Text = ""
'Text15.Text = ""
'Text16.Text = ""
'Text17.Text = ""

Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    If rs_stk_item_lgr!stk_item_lgr_name = Combo3.Text Or rs_stk_item_lgr!stk_item_lgr_alis = Combo3.Text Then
            
            Text1.Text = rs_stk_item_lgr!stk_item_lgr_name
            Text2.Text = rs_stk_item_lgr!stk_item_lgr_alis
            Text3.Text = rs_stk_item_lgr!stk_item_lgr_comp
            'Text4.Text = rs_stk_item_lgr!stk_item_lgr_grup
            
            Text5.Text = rs_stk_item_lgr!stk_item_lgr_unit
            Text6.Text = rs_stk_item_lgr!stk_item_lgr_fcvl
            Text7.Text = rs_stk_item_lgr!stk_item_lgr_disc

            Text8.Text = rs_stk_item_lgr!stk_item_lgr_qnt1
            
            Text9.Text = rs_stk_item_lgr!stk_item_lgr_qnt2
            
            Text10.Text = Format(rs_stk_item_lgr!stk_item_lgr_rat1, "0.00")
            Text11.Text = Format(rs_stk_item_lgr!stk_item_lgr_rat2, "0.00")
            
            Text12.Text = Format(rs_stk_item_lgr!stk_item_lgr_amt1, "0.00")
            Text13.Text = Format(rs_stk_item_lgr!stk_item_lgr_amt2, "0.00")
            
            Combo1.Text = rs_stk_item_lgr!stk_item_lgr_grup
            Combo2.Text = rs_stk_item_lgr!stk_item_lgr_unit
    
    Exit Sub
    End If

rs_stk_item_lgr.MoveNext
Loop
End Sub

Private Sub Command1_Click()

If Text1.Text = "" Or Text3.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox "You have to Enter all values first then....click this button...!!! Thank's"
Exit Sub
End If

selected_stock_item_name = Text1.Text
selected_stock_item_comp = Text3.Text
selected_stock_item_grup = Combo1.Text
selected_stock_item_unit = Combo2.Text
selected_stock_item_fval = Text6.Text
selected_stock_item_disc = Text7.Text
selected_stock_item_type = 1

'selected_stock_item_total_qnty = Text2.Text
'selected_stock_item_total_arat = Text2.Text
'selected_stock_item_total_amnt = Text2.Text

opening_stock_detail = "deatail_of_stock 1"
selected_stock_item = Text1.Text
If selected_stock_item = "" Then
    MsgBox "You have not selected any item...!! Please select any item first then enter the detail...!!!"
Exit Sub
End If
opg_stock_2_detail = 0
opg_stock_1_detail = 1
creat_stock_lgr_detail.Show
End Sub

Private Sub Command2_Click()

If Text1.Text = "" Or Text3.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox "You have to Enter all values first then....click this button...!!! Thank's"
Exit Sub
End If

selected_stock_item_name = Text1.Text
selected_stock_item_comp = Text3.Text
selected_stock_item_grup = Combo1.Text
selected_stock_item_unit = Combo2.Text
selected_stock_item_fval = Text6.Text
selected_stock_item_disc = Text7.Text
selected_stock_item_type = 2
'selected_stock_item_total_qnty = Text2.Text
'selected_stock_item_total_arat = Text2.Text
'selected_stock_item_total_amnt = Text2.Text

opening_stock_detail = "deatail_of_stock 2"
selected_stock_item = Text1.Text
If selected_stock_item = "" Then
    MsgBox "You have not selected any item...!! Please select any item first then enter the detail...!!!"
Exit Sub
End If

opg_stock_1_detail = 0
opg_stock_2_detail = 1
creat_stock_lgr_detail.Show

End Sub

Private Sub Form_Activate()
Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)

If opg_stock_1_detail = 2 Then
    Text8.Text = selected_stock_item_qnt1
    Text10.Text = Format(selected_stock_item_rat1, "0.00")
    Text12.Text = Format(selected_stock_item_amt1, "0.00")
ElseIf opg_stock_2_detail = 2 Then
    Text9.Text = selected_stock_item_qnt2
    Text11.Text = Format(selected_stock_item_rat2, "0.00")
    Text13.Text = Format(selected_stock_item_amt2, "0.00")
End If
End Sub

Private Sub Form_Load()
opg_stock_1_detail = 0
opg_stock_2_detail = 0
'selected_procedure = "Stock_item_ledger_edit"
'selected_procedure = "Stock_item_ledger_creat"
If selected_procedure = "Stock_item_ledger_edit" Then
    Label17.Visible = True
    Combo3.Visible = True
ElseIf selected_procedure = "Stock_item_ledger_creat" Then
    Label17.Visible = False
    Combo3.Visible = False
ElseIf selected_procedure = "Stock_item_ledger_display" Then
    Label17.Visible = True
    Combo3.Visible = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Visible = False
Text15.Visible = False
Text16.Visible = False
Text17.Visible = False
Text18.Visible = False
Combo1.Enabled = False
Combo2.Enabled = False
cmd_save.Enabled = False

End If

lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")

If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If

Call arrange_form_item
End Sub
Public Sub arrange_form_item()
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo4.Visible = False
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
Label15.Caption = ""
Label16.Caption = ""
Label17.Caption = ""
Label18.Caption = ""
Label19.Caption = ""

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
'Text14.Text = ""
'Text15.Text = ""
'Text16.Text = ""
'Text17.Text = ""
'Text18.Text = ""

Text1.FontSize = 12
Text2.FontSize = 12
Text3.FontSize = 12
Text4.FontSize = 12
Text5.FontSize = 12
Text6.FontSize = 12
Text7.FontSize = 12
Text8.FontSize = 12
Text9.FontSize = 12
Text10.FontSize = 12
Text11.FontSize = 12
Text12.FontSize = 12
Text13.FontSize = 12
'Text14.FontSize = 12
'Text15.FontSize = 12
'Text16.FontSize = 12
'Text17.FontSize = 12
'Text18.FontSize = 12

Text14.Visible = False
Text15.Visible = False
Text16.Visible = False
Text17.Visible = False
Text18.Visible = False

Label1.Caption = "Item Name"
Label2.Caption = "Alias"
Label3.Caption = "Company"
Label4.Caption = "Stock Group"
Label5.Caption = "Stock Unit"
Label6.Caption = "Item Face Value"
Label7.Caption = "item Discount %"
Label8.Caption = "Opening Stock 1"
Label9.Caption = "Opening Stock 2"
Label10.Caption = "Quantity 1"
Label11.Caption = "Quantity 2"
Label12.Caption = "Rate 1"
Label13.Caption = "Rate 2"
Label14.Caption = "Amount 1"
Label15.Caption = "Amount 2"
'Label16.Caption = "Cr/Dr"
Label17.Caption = "Select a Item"
'Label18.Caption = "Opening Balance 2"
'Label19.Caption = "Cr/Dr"
Combo1.Left = Text4.Left
Combo1.Top = Text4.Top
'Combo1.Height = Text12.Height
Combo1.Width = Text4.Width
Combo1.FontSize = 12
Combo2.Left = Text5.Left
Combo2.Top = Text5.Top
'Combo1.Height = Text12.Height
Combo2.Width = Text5.Width
Combo2.FontSize = 12
Call add_combo2_unit
Call add_combo1_main_grp
Combo3.Text = "Select a Item to edit"
End Sub
Public Sub add_combo2_unit()
Call open_database
Call open_rs_stk_item_unt
Do Until rs_stk_item_unt.EOF
    Combo2.AddItem rs_stk_item_unt!stk_item_unt_sybl & "(" & rs_stk_item_unt!stk_item_unt_name & ")"
rs_stk_item_unt.MoveNext
Loop
End Sub
Public Sub add_combo1_main_grp()
Call open_database
Call open_rs_stk_item_grp
 
Do Until rs_stk_item_grp.EOF
        Combo1.AddItem rs_stk_item_grp!stk_item_grp_name
        
        If rs_stk_item_grp!stk_item_grp_alis <> "" Then
            Combo1.AddItem rs_stk_item_grp!stk_item_grp_alis
        End If
        rs_stk_item_grp.MoveNext
Loop

Call SortList(Combo1, Val(0) \ 1, (Val(Combo1.ListCount) - 1) \ 1, Ascending)
Combo1.Text = "Select Group"

Call open_database
Call open_rs_stk_item_lgr
Do Until rs_stk_item_lgr.EOF
    Combo3.AddItem rs_stk_item_lgr!stk_item_lgr_name
    If rs_stk_item_lgr!stk_item_lgr_alis <> "" Then Combo3.AddItem rs_stk_item_lgr!stk_item_lgr_alis
rs_stk_item_lgr.MoveNext
Loop
Call SortList(Combo1, Val(0) \ 1, (Val(Combo1.ListCount) - 1) \ 1, Ascending)
Call SortList(Combo3, Val(0) \ 1, (Val(Combo3.ListCount) - 1) \ 1, Ascending)
End Sub

