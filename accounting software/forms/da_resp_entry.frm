VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form da_resp_entry 
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   12420
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Select one of the below"
      Height          =   615
      Left            =   360
      TabIndex        =   52
      Top             =   600
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "Suppler"
         Height          =   255
         Left            =   3840
         TabIndex        =   54
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Customer"
         Height          =   255
         Left            =   2400
         TabIndex        =   53
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1560
      TabIndex        =   50
      Text            =   "Select Customer...!!!"
      Top             =   1320
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DTPicker6 
      Height          =   495
      Left            =   17400
      TabIndex        =   46
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
      Format          =   107151361
      CurrentDate     =   40151
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   495
      Left            =   12600
      TabIndex        =   45
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
      Format          =   107151361
      CurrentDate     =   40151
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   495
      Left            =   7800
      TabIndex        =   44
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
      Format          =   107151361
      CurrentDate     =   40151
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   17400
      TabIndex        =   36
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Height          =   525
      Left            =   17400
      TabIndex        =   35
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   17400
      TabIndex        =   34
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   17400
      TabIndex        =   33
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   12600
      TabIndex        =   28
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      Height          =   525
      Left            =   12600
      TabIndex        =   27
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   12600
      TabIndex        =   26
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   12600
      TabIndex        =   25
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   7800
      TabIndex        =   23
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   7800
      TabIndex        =   20
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   7800
      TabIndex        =   18
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7800
      TabIndex        =   16
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command104 
      Caption         =   "Save"
      Height          =   375
      Left            =   19080
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text102 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "Type"
      Top             =   -500
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   -500
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   107151361
      CurrentDate     =   40149
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   4215
   End
   Begin VB.ComboBox Combo101 
      Height          =   315
      ItemData        =   "da_resp_entry.frx":0000
      Left            =   600
      List            =   "da_resp_entry.frx":0002
      TabIndex        =   6
      Text            =   "Select"
      Top             =   -500
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6000
      Left            =   360
      TabIndex        =   7
      Top             =   4440
      Width           =   19605
      _ExtentX        =   34581
      _ExtentY        =   10583
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   107151361
      CurrentDate     =   40141
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1560
      TabIndex        =   0
      Text            =   "Select Supplier...!!!"
      Top             =   1800
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   107151361
      CurrentDate     =   40141
   End
   Begin VB.Label Label24 
      Caption         =   "Supplier"
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
      TabIndex        =   51
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "Entry By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15720
      TabIndex        =   49
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Entry By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   48
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Entry By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   47
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   3135
      Left            =   6000
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Payment Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15840
      TabIndex        =   43
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmation Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   42
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Response Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   41
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label17 
      Caption         =   "Refrence No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15720
      TabIndex        =   40
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15720
      TabIndex        =   39
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15720
      TabIndex        =   38
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15720
      TabIndex        =   37
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Refrence No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   32
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   31
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   30
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   29
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Deactivated Cards Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   6720
      TabIndex        =   24
      Top             =   120
      Width           =   12375
   End
   Begin VB.Label Label8 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   22
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Refrence No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Period...,"
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
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
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
      Left            =   360
      TabIndex        =   14
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "To"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Customer"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6000
      Top             =   600
      Width           =   4575
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   10680
      Top             =   600
      Width           =   4575
   End
   Begin VB.Shape Shape3 
      Height          =   3135
      Left            =   10680
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   15360
      Top             =   600
      Width           =   4695
   End
   Begin VB.Shape Shape5 
      Height          =   3135
      Left            =   15360
      Top             =   1080
      Width           =   4695
   End
End
Attribute VB_Name = "da_resp_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'add customers to combo
Combo2.Clear
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup 'combo2.Text
selected_primary_group = ""
        Call open_rs_lgr_main_grp
        Do Until rs_lgr_main_grp.EOF
            If selected_group = rs_lgr_main_grp!lgr_main_grp_name Or selected_group = rs_lgr_main_grp!lgr_main_grp_alis Then
            selected_primary_group = rs_lgr_main_grp!lgr_main_grp_pgrp
            End If
            rs_lgr_main_grp.MoveNext
        Loop
        If selected_primary_group = "" Then
            Call open_rs_lgr_prim_grp
            If rs_lgr_prim_grp.RecordCount > 0 Then rs_lgr_prim_grp.MoveFirst
            Do Until rs_lgr_prim_grp.EOF
            If selected_group = rs_lgr_prim_grp!lgr_prim_grp_name Then
            selected_primary_group = rs_lgr_prim_grp!lgr_prim_grp_name
            End If
            rs_lgr_prim_grp.MoveNext
            Loop
        End If
        If LCase(selected_primary_group) = LCase("Sundry Debtors") Then ' if the created ledger is a debtor then
            Combo2.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        End If
rs_lgr_main_dtl.MoveNext
Loop
Combo2.Text = "Select Customer..,"
'=============================
Option2.Value = True
Combo2.Enabled = False
Me.Caption = "Ajay patel's card Deactivation...!!!  " & user_name
Grid1.RowHeightMin = 400
'Call add_item_in_combo1 'for supplier
'=============================
Combo1.Clear
Call open_database
Call open_rs_lgr_main_dtl
Do Until rs_lgr_main_dtl.EOF
selected_group = rs_lgr_main_dtl!lgr_main_dtl_grup 'combo1.Text
selected_primary_group = ""
        Call open_rs_lgr_main_grp
        Do Until rs_lgr_main_grp.EOF
            If selected_group = rs_lgr_main_grp!lgr_main_grp_name Or selected_group = rs_lgr_main_grp!lgr_main_grp_alis Then
            selected_primary_group = rs_lgr_main_grp!lgr_main_grp_pgrp
            End If
            rs_lgr_main_grp.MoveNext
        Loop
        If selected_primary_group = "" Then
            Call open_rs_lgr_prim_grp
            If rs_lgr_prim_grp.RecordCount > 0 Then rs_lgr_prim_grp.MoveFirst
            Do Until rs_lgr_prim_grp.EOF
            If selected_group = rs_lgr_prim_grp!lgr_prim_grp_name Then
            selected_primary_group = rs_lgr_prim_grp!lgr_prim_grp_name
            End If
            rs_lgr_prim_grp.MoveNext
            Loop
        End If
        If LCase(selected_primary_group) = LCase("Sundry creditors") Then ' if the created ledger is a debtor then
            Combo1.AddItem rs_lgr_main_dtl!lgr_main_dtl_name
        End If
rs_lgr_main_dtl.MoveNext
Loop
Combo1.Text = "Select Customer..,"
'Call add_item_in_combo101 'for select a column
'for select a period
Combo4.AddItem "This Month"
Combo4.AddItem "This Week"
Combo4.AddItem "Last Month"
Combo4.AddItem "Last Week"
Call set_grid1_data
Combo4.Text = "Last Month"
Call search_a_period
'Command104.Left = Grid1.CellLeft(21, 0)
End Sub
Public Sub search_record_and_save_to_dpdb2_res()
'MsgBox Grid1.TextMatrix(Grid1.Row, 1)
'MsgBox Grid1.TextMatrix(Grid1.Row, 2)
'MsgBox Grid1.TextMatrix(Grid1.Row, 3)
'MsgBox Grid1.TextMatrix(Grid1.Row, 4)
Dim aa
Close All
'Call open_dpdb2
aa = Grid1.Row
Do Until rs_dap_rspn_dtl.EOF
If Grid1.TextMatrix(Grid1.Row, 1) = rs_dap_rspn_dtl!dap_main_rsp_dt And _
        Grid1.TextMatrix(Grid1.Row, 2) = rs_dap_rspn_dtl!dap_main_rsp_rf And _
        Grid1.TextMatrix(Grid1.Row, 3) = rs_dap_rspn_dtl!db2_cust_nm And _
        Grid1.TextMatrix(Grid1.Row, 4) = rs_dap_rspn_dtl!db2_supl_nm Then
        If Grid1.TextMatrix(aa, 6) <> Null Or Grid1.TextMatrix(aa, 6) <> "" Then
            rs_dap_rspn_dtl!dap_main_rsp_rsdt = Grid1.TextMatrix(aa, 6)
        Else:
         '   rs_dap_rspn_dtl!dap_main_rsp_rsdt = ""
        End If
        If Grid1.TextMatrix(aa, 7) <> Null Or Grid1.TextMatrix(aa, 7) <> "" Then
         rs_dap_rspn_dtl!dap_main_rsp_rsrf = Grid1.TextMatrix(aa, 7)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_rsrf = ""
        End If
        If Grid1.TextMatrix(aa, 8) <> Null Or Grid1.TextMatrix(aa, 8) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_rstp = Grid1.TextMatrix(aa, 8)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_rstp = ""
        End If
        If Grid1.TextMatrix(aa, 9) <> Null Or Grid1.TextMatrix(aa, 9) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_rsam = Grid1.TextMatrix(aa, 9)
        'Else:
        '    rs_dap_rspn_dtl!dap_main_rsp_rsam = ""
        End If
        If Grid1.TextMatrix(aa, 10) <> Null Or Grid1.TextMatrix(aa, 10) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_rsus = Grid1.TextMatrix(aa, 10)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_rsus = ""
        End If
        If Grid1.TextMatrix(aa, 11) <> Null Or Grid1.TextMatrix(aa, 11) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_cndt = Grid1.TextMatrix(aa, 11)
        'Else:
        '    rs_dap_rspn_dtl!dap_main_rsp_cndt = ""
        End If
        If Grid1.TextMatrix(aa, 12) <> Null Or Grid1.TextMatrix(aa, 12) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_cnrf = Grid1.TextMatrix(aa, 12)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_cnrf = ""
        End If
        If Grid1.TextMatrix(aa, 13) <> Null Or Grid1.TextMatrix(aa, 13) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_cntp = Grid1.TextMatrix(aa, 13)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_cntp = ""
        End If
        If Grid1.TextMatrix(aa, 14) <> Null Or Grid1.TextMatrix(aa, 14) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_cnam = Grid1.TextMatrix(aa, 14)
        'Else:
        '    rs_dap_rspn_dtl!dap_main_rsp_cnam = ""
        End If
        If Grid1.TextMatrix(aa, 15) <> Null Or Grid1.TextMatrix(aa, 15) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_cnus = Grid1.TextMatrix(aa, 15)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_cnus = ""
        End If
        
        If Grid1.TextMatrix(aa, 16) <> Null Or Grid1.TextMatrix(aa, 16) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_pmdt = Grid1.TextMatrix(aa, 16)
        'Else:
        '    rs_dap_rspn_dtl!dap_main_rsp_pmdt = ""
        End If
        If Grid1.TextMatrix(aa, 17) <> Null Or Grid1.TextMatrix(aa, 17) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_pmrf = Grid1.TextMatrix(aa, 17)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_pmrf = ""
        End If
        
        If Grid1.TextMatrix(aa, 18) <> Null Or Grid1.TextMatrix(aa, 18) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_pmtp = Grid1.TextMatrix(aa, 18)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_pmtp = ""
        End If
        
        If Grid1.TextMatrix(aa, 19) <> Null Or Grid1.TextMatrix(aa, 19) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_pmam = Grid1.TextMatrix(aa, 19)
        'Else:
        '    rs_dap_rspn_dtl!dap_main_rsp_pmam = ""
        End If
        If Grid1.TextMatrix(aa, 20) <> Null Or Grid1.TextMatrix(aa, 20) <> "" Then
        rs_dap_rspn_dtl!dap_main_rsp_pmus = Grid1.TextMatrix(aa, 20)
        Else:
            rs_dap_rspn_dtl!dap_main_rsp_pmus = ""
        End If
        rs_dap_rspn_dtl.UpdateBatch
'        Exit Do
End If
rs_dap_rspn_dtl.MoveNext
Loop
        

End Sub
Private Sub Command104_Click()
Call search_record_and_save_to_dpdb2_res
End Sub

Private Sub Command2_Click()
If user_name = "" Then user_name = "not selected"
'selecting customer name through combo & date through date pick button

selected_str_date = DTPicker1.Value
selected_end_date = DTPicker2.Value

'If dpdb.State = 1 Then dpdb.Close
'find data of supplier Response Data from x to y date
'dpdb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\dact_db.mdb;"
'dpdb.Open


Call open_database
rs_dap_rspn_dtl.CursorLocation = adUseClient
If Option2.Value = True Then
    selected_supl_name = Combo1.Text
    rs_dap_rspn_dtl.Open "SELECT * FROM dap_rspn_dtl WHERE dap_main_rsp_trdt>= " & selected_str_date & " and dap_main_rsp_spnm = '" & selected_supl_name & "'and dap_main_rsp_trdt <=#" & selected_end_date & "#", db_co, adOpenDynamic, adLockOptimistic
ElseIf Option1.Value = True Then
    selected_cust_nm = Combo2.Text
    rs_dap_rspn_dtl.Open "SELECT * FROM dap_rspn_dtl WHERE dap_main_rsp_trdt>= " & selected_str_date & " and dap_main_rsp_csnm = '" & selected_cust_nm & "'and dap_main_rsp_trdt <=#" & selected_end_date & "#", db_co, adOpenDynamic, adLockOptimistic
End If

With dact_repo_resp_conf_pmnt.Sections("section2").Controls
    .item("label13").Caption = temp_ref_no
End With

Set dact_repo_resp_conf_pmnt.DataSource = rs_dap_rspn_dtl.DataSource
dact_repo_resp_conf_pmnt.Show

End Sub

Private Sub Command1_Click()
Call set_grid1_data
'find which user is operating the computer & doing the work
If user_name = "" Then user_name = "not selected"
'selecting customer name through combo & date through date pick button
selected_str_date = DTPicker1.Value
selected_end_date = DTPicker2.Value

Call open_database
rs_dap_rspn_dtl.CursorLocation = adUseClient
If Option2.Value = True Then
    selected_supl_name = Combo1.Text
    rs_dap_rspn_dtl.Open "SELECT * FROM dap_rspn_dtl WHERE dap_main_rsp_trdt>= " & selected_str_date & " and dap_main_rsp_spnm = '" & selected_supl_name & "'and dap_main_rsp_trdt <=#" & selected_end_date & "#", db_co, adOpenDynamic, adLockOptimistic
ElseIf Option1.Value = True Then
    selected_cust_nm = Combo2.Text
    rs_dap_rspn_dtl.Open "SELECT * FROM dap_rspn_dtl WHERE dap_main_rsp_trdt>= " & selected_str_date & " and dap_main_rsp_csnm = '" & selected_cust_nm & "'and dap_main_rsp_trdt <=#" & selected_end_date & "#", db_co, adOpenDynamic, adLockOptimistic
End If
Dim aa
aa = 1
Do Until rs_dap_rspn_dtl.EOF
        Grid1.AddItem aa
        Grid1.TextMatrix(aa, 1) = rs_dap_rspn_dtl!dap_main_rsp_trdt
        Grid1.TextMatrix(aa, 2) = rs_dap_rspn_dtl!dap_main_rsp_trrf
        Grid1.TextMatrix(aa, 3) = rs_dap_rspn_dtl!dap_main_rsp_csnm
        Grid1.TextMatrix(aa, 4) = rs_dap_rspn_dtl!dap_main_rsp_spnm
        Grid1.TextMatrix(aa, 5) = rs_dap_rspn_dtl!dap_main_rsp_trus
        If rs_dap_rspn_dtl!dap_main_rsp_rsdt <> Null Or rs_dap_rspn_dtl!dap_main_rsp_rsdt <> "" Then
            Grid1.TextMatrix(aa, 6) = rs_dap_rspn_dtl!dap_main_rsp_rsdt
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_rsrf <> Null Or rs_dap_rspn_dtl!dap_main_rsp_rsrf <> "" Then
        Grid1.TextMatrix(aa, 7) = rs_dap_rspn_dtl!dap_main_rsp_rsrf
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_rstp <> Null Or rs_dap_rspn_dtl!dap_main_rsp_rstp <> "" Then
        Grid1.TextMatrix(aa, 8) = rs_dap_rspn_dtl!dap_main_rsp_rstp
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_rsam <> Null Or rs_dap_rspn_dtl!dap_main_rsp_rsam <> "" Then
        Grid1.TextMatrix(aa, 9) = rs_dap_rspn_dtl!dap_main_rsp_rsam
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_rsus <> Null Or rs_dap_rspn_dtl!dap_main_rsp_rsus <> "" Then
        Grid1.TextMatrix(aa, 10) = rs_dap_rspn_dtl!dap_main_rsp_rsus
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_cndt <> Null Or rs_dap_rspn_dtl!dap_main_rsp_cndt <> "" Then
        Grid1.TextMatrix(aa, 11) = rs_dap_rspn_dtl!dap_main_rsp_cndt
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_cnrf <> Null Or rs_dap_rspn_dtl!dap_main_rsp_cnrf <> "" Then
        Grid1.TextMatrix(aa, 12) = rs_dap_rspn_dtl!dap_main_rsp_cnrf
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_cntp <> Null Or rs_dap_rspn_dtl!dap_main_rsp_cntp <> "" Then
        Grid1.TextMatrix(aa, 13) = rs_dap_rspn_dtl!dap_main_rsp_cntp
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_cnam <> Null Or rs_dap_rspn_dtl!dap_main_rsp_cnam <> "" Then
        Grid1.TextMatrix(aa, 14) = rs_dap_rspn_dtl!dap_main_rsp_cnam
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_cnus <> Null Or rs_dap_rspn_dtl!dap_main_rsp_cnus <> "" Then
        Grid1.TextMatrix(aa, 15) = rs_dap_rspn_dtl!dap_main_rsp_cnus
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_pmdt <> Null Or rs_dap_rspn_dtl!dap_main_rsp_pmdt <> "" Then
        Grid1.TextMatrix(aa, 16) = rs_dap_rspn_dtl!dap_main_rsp_pmdt
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_pmrf <> Null Or rs_dap_rspn_dtl!dap_main_rsp_pmrf <> "" Then
        Grid1.TextMatrix(aa, 17) = rs_dap_rspn_dtl!dap_main_rsp_pmrf
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_pmtp <> Null Or rs_dap_rspn_dtl!dap_main_rsp_pmtp <> "" Then
        Grid1.TextMatrix(aa, 18) = rs_dap_rspn_dtl!dap_main_rsp_pmtp
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_pmam <> Null Or rs_dap_rspn_dtl!dap_main_rsp_pmam <> "" Then
        Grid1.TextMatrix(aa, 19) = rs_dap_rspn_dtl!dap_main_rsp_pmam
        End If
        If rs_dap_rspn_dtl!dap_main_rsp_pmus <> Null Or rs_dap_rspn_dtl!dap_main_rsp_pmus <> "" Then
        Grid1.TextMatrix(aa, 20) = rs_dap_rspn_dtl!dap_main_rsp_pmus
        End If
        aa = aa + 1
rs_dap_rspn_dtl.MoveNext
Loop

End Sub
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Grid1_Click()
Call show_data
'when click on such column how to active a combo
If Grid1.Col > 5 And Grid1.Col < 21 Then
        Command104.Height = Grid1.CellHeight
        Command104.Top = Grid1.CellTop + Grid1.Top
        Command104.Visible = True
End If
If p_grid_col <> Grid1.Col Or p_grid_row <> Grid1.Row Then
Text102.Text = ""
End If
    If Grid1.Col = 2 Then      ' Position and size the ComboBox, then show it.
        'Combo101.Height = Grid1.CellHeight
        'Combo101.Width = Grid1.CellWidth
        'Combo101.Left = Grid1.CellLeft + Grid1.Left
        'Combo101.Top = Grid1.CellTop + Grid1.Top
        'Combo101.Text = Grid1.Text
        'Combo101.Visible = True
    ElseIf Grid1.Col = 6 Or Grid1.Col = 11 Or Grid1.Col = 16 Then
        DTPicker3.Height = Grid1.CellHeight
        DTPicker3.Width = Grid1.CellWidth
        DTPicker3.Left = Grid1.CellLeft + Grid1.Left
        DTPicker3.Top = Grid1.CellTop + Grid1.Top
        DTPicker3.Value = Date
        DTPicker3.Visible = True
    ElseIf Grid1.Col = 7 Or Grid1.Col = 8 Or Grid1.Col = 9 _
        Or Grid1.Col = 12 Or Grid1.Col = 13 Or Grid1.Col = 14 _
        Or Grid1.Col = 17 Or Grid1.Col = 18 Or Grid1.Col = 19 Then
            Text102.Height = Grid1.CellHeight
            Text102.Width = Grid1.CellWidth
            Text102.Left = Grid1.CellLeft + Grid1.Left
            Text102.Top = Grid1.CellTop + Grid1.Top
            'Text102.Value = Date
            Text102.Visible = True
            'ElseIf Grid1.Col = 6 Or Grid1.Col = 11 Or Grid1.Col = 16 Then
    End If
    
End Sub

Public Sub search_a_period()
Dim today_day As Integer
Dim today_weekday As Integer

today_weekday = Weekday(Now)
today_day = Day(Now) - 1

If Combo4.Text = "This Week" Then
    DTPicker1.Value = Date - (today_weekday + 1)
    DTPicker2.Value = Date
ElseIf Combo4.Text = "This Month" Then
    DTPicker1.Value = Date - today_day
    DTPicker2.Value = Date
ElseIf Combo4.Text = "Last Month" Then
    If Month(Now) = 1 Then
        DTPicker1.Value = Day(Now) - today_day & "/12/" & Year(Now) - 1
    Else
    DTPicker1.Value = Day(Now) - today_day & "/" & Month(Now) - 1 & "/" & Year(Now)
    End If
    DTPicker2.Value = Date - (today_day + 1)
ElseIf Combo4.Text = "Last Week" Then
    DTPicker1.Value = Date - (today_weekday + 5)
    DTPicker2.Value = Date - (today_weekday - 1)
End If

End Sub
Private Sub DTPicker3_Change()
    If Grid1.Col = 6 Or Grid1.Col = 11 Or Grid1.Col = 16 Then
        Dim x_date As Date
        x_date = DTPicker3.Value
      Grid1.Text = x_date
      
                If Grid1.Col = 6 Then
                            Grid1.TextMatrix(Grid1.Row, 10) = user_name
                ElseIf Grid1.Col = 11 Then
                            Grid1.TextMatrix(Grid1.Row, 15) = user_name
                ElseIf Grid1.Col = 16 Then
                            Grid1.TextMatrix(Grid1.Row, 20) = user_name
                End If
                DTPicker3.Visible = False
    End If
End Sub


Private Sub Option1_Click()
Combo1.Enabled = False
Combo2.Enabled = True
End Sub

Private Sub Option2_Click()
Combo1.Enabled = True
Combo2.Enabled = False
End Sub

Private Sub Text102_Change()

    If Grid1.Col = 7 Or Grid1.Col = 8 Or Grid1.Col = 9 _
        Or Grid1.Col = 12 Or Grid1.Col = 13 Or Grid1.Col = 14 _
        Or Grid1.Col = 17 Or Grid1.Col = 18 Or Grid1.Col = 19 Then
                Grid1.Text = Text102.Text
                'Text102.Visible = False
    End If

p_grid_col = Grid1.Col
p_grid_row = Grid1.Row

End Sub
Private Sub combo101_Click()
    If Grid1.Col = 2 Then
      Grid1.Text = Combo101.Text
      Combo101.Visible = False
    End If
End Sub

Private Sub Combo4_Click()
Call search_a_period
End Sub
Public Sub add_item_in_combo101()
'set the combo for click
Combo101.AddItem "1"
Combo101.AddItem "2"
Combo101.AddItem "3"
End Sub
Public Sub add_item_in_combo4()
Combo4.AddItem "This Month"
Combo4.AddItem "This Week"
Combo4.AddItem "Last Month"
Combo4.AddItem "Last Week"
End Sub
Public Sub add_item_in_combo1()
'add customer from ledger
'Call open_dpdb6
Do Until dpdb6_rs.EOF
    'If dpdb5_rs!lgr_acnt_grup = "customer" Then Combo1.AddItem dpdb5_rs!lgr_acnt_name
    Combo1.AddItem dpdb6_rs!supl_name
    dpdb6_rs.MoveNext
Loop
End Sub

Public Sub set_grid1_data()
'set data grid
Grid1.Clear
Grid1.Rows = 1
Grid1.Cols = 22

Grid1.TextMatrix(0, 1) = "DA-Date"
Grid1.TextMatrix(0, 2) = "DA-Ref"
Grid1.TextMatrix(0, 3) = "Customer"
Grid1.TextMatrix(0, 4) = "Supplier"
Grid1.TextMatrix(0, 5) = "DA-By"

Grid1.TextMatrix(0, 6) = "Resp-Date"
Grid1.TextMatrix(0, 7) = "Resp-Ref"
Grid1.TextMatrix(0, 8) = "Resp-Type"
Grid1.TextMatrix(0, 9) = "Resp-Amt"
Grid1.TextMatrix(0, 10) = "Resp-by"

Grid1.TextMatrix(0, 11) = "conf-Date"
Grid1.TextMatrix(0, 12) = "conf-Ref"
Grid1.TextMatrix(0, 13) = "conf-Type"
Grid1.TextMatrix(0, 14) = "conf-Amt"
Grid1.TextMatrix(0, 15) = "conf-by"

Grid1.TextMatrix(0, 16) = "Pay-Date"
Grid1.TextMatrix(0, 17) = "pay-Ref"
Grid1.TextMatrix(0, 18) = "pay-Type"
Grid1.TextMatrix(0, 19) = "pay-Amt"
Grid1.TextMatrix(0, 20) = "pay-by"
Grid1.TextMatrix(0, 21) = "Save"

Grid1.ColWidth(0) = 400

Grid1.ColWidth(1) = 1000
Grid1.ColWidth(2) = 1700
Grid1.ColWidth(3) = 2000
Grid1.ColWidth(4) = 800
Grid1.ColWidth(5) = 600

Grid1.ColWidth(6) = 1000
Grid1.ColWidth(7) = 800
Grid1.ColWidth(8) = 1000
Grid1.ColWidth(9) = 1000
Grid1.ColWidth(10) = 600

Grid1.ColWidth(11) = 1000
Grid1.ColWidth(12) = 1000
Grid1.ColWidth(13) = 1000
Grid1.ColWidth(14) = 1000
Grid1.ColWidth(15) = 600

Grid1.ColWidth(16) = 1000
Grid1.ColWidth(17) = 1000
Grid1.ColWidth(18) = 1000
Grid1.ColWidth(19) = 750
Grid1.ColWidth(20) = 600
Grid1.ColWidth(21) = 800

Command104.Width = Grid1.ColWidth(21)
Command104.Visible = False
Call all_disable
End Sub

Public Sub all_disable()

'DTPicker4.Value = Null

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

DTPicker4.Enabled = False
DTPicker5.Enabled = False
DTPicker6.Enabled = False

DTPicker4.Visible = False
DTPicker5.Visible = False
DTPicker6.Visible = False
End Sub
Public Sub show_data()
If Grid1.TextMatrix(Grid1.Row, 6) <> Null Or Grid1.TextMatrix(Grid1.Row, 6) <> "" Then
        DTPicker4.Value = Grid1.TextMatrix(Grid1.Row, 6)
        DTPicker4.Visible = True
End If
If Grid1.TextMatrix(Grid1.Row, 7) <> Null Or Grid1.TextMatrix(Grid1.Row, 7) <> "" Then Text1.Text = Grid1.TextMatrix(Grid1.Row, 7)
If Grid1.TextMatrix(Grid1.Row, 8) <> Null Or Grid1.TextMatrix(Grid1.Row, 8) <> "" Then Text2.Text = Grid1.TextMatrix(Grid1.Row, 8)
If Grid1.TextMatrix(Grid1.Row, 9) <> Null Or Grid1.TextMatrix(Grid1.Row, 9) <> "" Then Text3.Text = Grid1.TextMatrix(Grid1.Row, 9)
If Grid1.TextMatrix(Grid1.Row, 10) <> Null Or Grid1.TextMatrix(Grid1.Row, 10) <> "" Then Text4.Text = Grid1.TextMatrix(Grid1.Row, 10)
If Grid1.TextMatrix(Grid1.Row, 11) <> Null Or Grid1.TextMatrix(Grid1.Row, 11) <> "" Then
        DTPicker5.Value = Grid1.TextMatrix(Grid1.Row, 11)
        DTPicker5.Visible = True
End If
If Grid1.TextMatrix(Grid1.Row, 12) <> Null Or Grid1.TextMatrix(Grid1.Row, 12) <> "" Then Text5.Text = Grid1.TextMatrix(Grid1.Row, 12)
If Grid1.TextMatrix(Grid1.Row, 13) <> Null Or Grid1.TextMatrix(Grid1.Row, 13) <> "" Then Text6.Text = Grid1.TextMatrix(Grid1.Row, 13)
If Grid1.TextMatrix(Grid1.Row, 14) <> Null Or Grid1.TextMatrix(Grid1.Row, 14) <> "" Then Text7.Text = Grid1.TextMatrix(Grid1.Row, 14)
If Grid1.TextMatrix(Grid1.Row, 15) <> Null Or Grid1.TextMatrix(Grid1.Row, 15) <> "" Then Text8.Text = Grid1.TextMatrix(Grid1.Row, 15)
If Grid1.TextMatrix(Grid1.Row, 16) <> Null Or Grid1.TextMatrix(Grid1.Row, 16) <> "" Then
    DTPicker6.Value = Grid1.TextMatrix(Grid1.Row, 16)
    DTPicker6.Visible = True
End If
If Grid1.TextMatrix(Grid1.Row, 17) <> Null Or Grid1.TextMatrix(Grid1.Row, 17) <> "" Then Text9.Text = Grid1.TextMatrix(Grid1.Row, 17)
If Grid1.TextMatrix(Grid1.Row, 18) <> Null Or Grid1.TextMatrix(Grid1.Row, 18) <> "" Then Text10.Text = Grid1.TextMatrix(Grid1.Row, 18)
If Grid1.TextMatrix(Grid1.Row, 19) <> Null Or Grid1.TextMatrix(Grid1.Row, 19) <> "" Then Text11.Text = Grid1.TextMatrix(Grid1.Row, 19)
If Grid1.TextMatrix(Grid1.Row, 20) <> Null Or Grid1.TextMatrix(Grid1.Row, 20) <> "" Then Text12.Text = Grid1.TextMatrix(Grid1.Row, 20)
End Sub
