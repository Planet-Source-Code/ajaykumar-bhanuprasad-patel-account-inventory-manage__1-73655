VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Creat_st_unt 
   Caption         =   "Creat Group"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
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
      Left            =   3120
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   4680
      Width           =   5535
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      TabIndex        =   12
      Text            =   "Combo3"
      Top             =   3240
      Width           =   5535
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3120
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   4200
      Width           =   5500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3720
      Width           =   5500
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6210
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Unit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   8415
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
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   2415
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
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   7095
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
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   7335
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
      TabIndex        =   5
      Top             =   4200
      Width           =   2655
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
      TabIndex        =   4
      Top             =   3720
      Width           =   2535
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
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8640
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   0
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "Creat_st_unt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Caption = selected_company & ".../" & selected_procedure & ".../" & UCase(selected_user)

'selected_procedure = "stock_unit_display"
'selected_procedure = "stock_unit_edit"
'selected_procedure = "stock_unit_creat"

If selected_procedure = "stock_unit_edit" Then
    Label5.Visible = True
    Combo3.Visible = True
ElseIf selected_procedure = "stock_unit_creat" Then
    Label5.Visible = False
    Combo3.Visible = False
ElseIf selected_procedure = "stock_unit_display" Then
    Label5.Visible = True
    Combo3.Visible = True
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
End If
lbl_name.Caption = co_name
lbl_add.Caption = selected_companies_add1 & ", " & selected_companies_add2 & ", " & selected_companies_pincode & ", " & selected_companies_city & ", " & selected_companies_country
'Image1.Picture = LoadPicture(App.Path & "\icon\pic1.jpg")
If selected_path = "" Or selected_path = Null Then
    selected_path = App.Path & "\data\1000\co.mdb;"
End If
Call arrange_form
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub
Private Sub cmd_save_Click()
Dim selected_stk_group_alias
'check the data                 = if error      message
If Val(Text3.Text) < 0 Or Val(Text3.Text) >= 10 Then
MsgBox "There are Wrong decimal place you entered...!!!"
Exit Sub
End If

If Text2.Text = "" Or Text1.Text = "" Then
MsgBox "There is something is empty....you entered...!!!"
Exit Sub
End If
'check for duplicate Data
If selected_procedure = "stock_unit_edit" Then
            Call open_database
            Call open_rs_stk_item_unt
            Do Until rs_stk_item_unt.EOF
                If rs_stk_item_unt!stk_item_unt_name = Text1.Text Then
                        MsgBox "This Unit is already exist...!!!"
                        Exit Sub
                End If
                rs_stk_item_unt.MoveNext
            Loop
 Call open_database
 Call open_rs_stk_item_unt
 
 Do Until rs_stk_item_unt.EOF
        If Combo3.Text = rs_stk_item_unt!stk_item_unt_name Then
            rs_stk_item_unt!stk_item_unt_name = Text1.Text
            rs_stk_item_unt!stk_item_unt_sybl = Text2.Text
            rs_stk_item_unt!stk_item_unt_dcml = Text3.Text
            'rs_stk_item_unt!stk_item_unt_pgrp = Combo2.Text
            rs_stk_item_unt.UpdateBatch
        End If
        rs_stk_item_unt.MoveNext
 Loop
ElseIf selected_procedure = "stock_unit_creat" Then
            'open_file & find the data      = if available  message
            Call open_database
            Call open_rs_stk_item_unt
            Do Until rs_stk_item_unt.EOF
                If rs_stk_item_unt!stk_item_unt_name = Text1.Text Then
                        MsgBox "This Unit is already exist...!!!"
                        Exit Sub
                End If
                rs_stk_item_unt.MoveNext
            Loop
            'open_file to save a file   'save a record to file
            Call open_database
            Call open_rs_stk_item_unt
            rs_stk_item_unt.AddNew
                rs_stk_item_unt!stk_item_unt_name = Text1.Text
                rs_stk_item_unt!stk_item_unt_sybl = Text2.Text
                rs_stk_item_unt!stk_item_unt_dcml = Text3.Text
                'rs_stk_item_unt!stk_item_unt_pgrp = Combo2.Text
            rs_stk_item_unt.UpdateBatch
            rs_stk_item_unt.Close
End If
Call arrange_form
End Sub
Private Sub Combo3_Click()
 Call open_database
 Call open_rs_stk_item_unt
 Do Until rs_stk_item_unt.EOF
        If Combo3.Text = rs_stk_item_unt!stk_item_unt_name Then
            Text1.Text = rs_stk_item_unt!stk_item_unt_name
            Text2.Text = rs_stk_item_unt!stk_item_unt_sybl
            Text3.Text = rs_stk_item_unt!stk_item_unt_dcml
        End If
        rs_stk_item_unt.MoveNext
 Loop
End Sub
Public Sub arrange_form()
Combo3.Clear
Label1.Caption = "Unit Name"
Label2.Caption = "Unit Symbol"
Label3.Caption = "Decimal"
Label5.Caption = "Select Unit"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo3.Text = ""
Call add_combo3_main_grp
End Sub
Public Sub add_combo3_main_grp()
 Call open_database
 Call open_rs_stk_item_unt
 Do Until rs_stk_item_unt.EOF
        Combo3.AddItem rs_stk_item_unt!stk_item_unt_name
        rs_stk_item_unt.MoveNext
 Loop
End Sub
