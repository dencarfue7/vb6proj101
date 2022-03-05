VERSION 5.00
Begin VB.Form Main_Form 
   Caption         =   "Singly Reinforce Beam Design"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_clearFields 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear Fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton btn_viewStirrups 
      BackColor       =   &H00FFFFC0&
      Caption         =   "View Stirrups Design"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   8160
      Width           =   2655
   End
   Begin VB.CommandButton btn_calculate 
      BackColor       =   &H0080FF80&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   8160
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Steel Reinforcements : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Index           =   5
      Left            =   9480
      TabIndex        =   63
      Top             =   4200
      Width           =   4335
      Begin VB.CommandButton btn_details 
         Caption         =   "Details"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   70
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txt_srAtSupport 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   405
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txt_srAtMidspan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   405
         Left            =   2040
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txt_wr1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   405
         Left            =   2040
         TabIndex        =   2
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txt_wr2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   405
         Left            =   2040
         TabIndex        =   1
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Web Reinforcements: "
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   82
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "At Midspan: "
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   81
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "At Support: "
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Beam Adequacy : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3735
      Index           =   4
      Left            =   4200
      TabIndex        =   10
      Top             =   4200
      Width           =   5055
      Begin VB.TextBox txt_dreq 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   31
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txt_hreq 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   32
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txt_rn 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   30
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txt_w 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   29
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txt_pbal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   28
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txt_pmax 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lbl_adeq_ok 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK Label"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   3240
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   255
         Left            =   4320
         TabIndex        =   79
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   255
         Left            =   4320
         TabIndex        =   78
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "MPa"
         Height          =   255
         Left            =   4320
         TabIndex        =   77
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "h req'd :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "d req'd :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Rn :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "w :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Pbal :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Pmax :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Design Data : "
      Height          =   3735
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   3615
      Begin VB.TextBox txt_shear 
         Height          =   375
         Left            =   960
         TabIndex        =   26
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txt_mumid 
         Height          =   405
         Left            =   960
         TabIndex        =   25
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txt_musup 
         Height          =   375
         Left            =   960
         TabIndex        =   24
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label56 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   600
         TabIndex        =   93
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label55 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   840
         TabIndex        =   92
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label54 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   840
         TabIndex        =   91
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label34 
         Caption         =   "kN"
         Height          =   255
         Left            =   3000
         TabIndex        =   69
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label33 
         Caption         =   "kN-m"
         Height          =   255
         Left            =   3000
         TabIndex        =   68
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label32 
         Caption         =   "kN-m"
         Height          =   255
         Left            =   3000
         TabIndex        =   67
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label31 
         Caption         =   "Shear, Vu:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   62
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "MUmid:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "MUsup:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "At Midspan: "
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   59
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label27 
         Caption         =   "At Supports: "
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Beam Geometry : "
      Height          =   2895
      Index           =   2
      Left            =   9480
      TabIndex        =   8
      Top             =   960
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "Compute"
         Height          =   375
         Left            =   1800
         TabIndex        =   96
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txt_effectiveDepth 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txt_length 
         Height          =   405
         Left            =   1800
         TabIndex        =   22
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txt_height 
         Height          =   405
         Left            =   1800
         TabIndex        =   21
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txt_base 
         Height          =   405
         Left            =   1800
         TabIndex        =   20
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label53 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   480
         TabIndex        =   90
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label52 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1200
         TabIndex        =   89
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label51 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   960
         TabIndex        =   88
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label26 
         Caption         =   "mm"
         Height          =   255
         Left            =   3840
         TabIndex        =   57
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "m"
         Height          =   255
         Left            =   3840
         TabIndex        =   56
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "mm"
         Height          =   255
         Left            =   3840
         TabIndex        =   55
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "Effective Depth, d:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   54
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "Span Length, L:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   53
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Height, h:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "mm"
         Height          =   255
         Left            =   3840
         TabIndex        =   51
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Base, b:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Rebars Geometry : "
      Height          =   3135
      Index           =   1
      Left            =   4200
      TabIndex        =   7
      Top             =   960
      Width           =   5055
      Begin VB.CommandButton Command2 
         Caption         =   "Compute"
         Height          =   375
         Left            =   2280
         TabIndex        =   95
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txt_concreteCover 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   19
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txt_nominalAreaAnt 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txt_stirrups 
         Height          =   405
         Left            =   2280
         TabIndex        =   17
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txt_nominalAreaAn 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txt_mainbar 
         Height          =   405
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label47 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1560
         TabIndex        =   84
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label13 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1680
         TabIndex        =   83
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label18 
         Caption         =   "mm"
         Height          =   255
         Left            =   4320
         TabIndex        =   49
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "mmsq."
         Height          =   255
         Left            =   4320
         TabIndex        =   48
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "mmÿ"
         Height          =   255
         Left            =   4320
         TabIndex        =   47
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "mmsq."
         Height          =   255
         Left            =   4320
         TabIndex        =   46
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "mmÿ"
         Height          =   255
         Left            =   4320
         TabIndex        =   45
         Top             =   360
         Width           =   495
      End
      Begin VB.Label label 
         Caption         =   "Concrete Cover, Cc:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Nominal Area, Ant:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Stirrups, dst:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Nominal Area, An:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Main Bars, db:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Material Properties : "
      Height          =   2895
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "Compute"
         Height          =   375
         Left            =   960
         TabIndex        =   94
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txt_b1 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txt_fyt 
         Height          =   405
         Left            =   960
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txt_fy 
         Height          =   405
         Left            =   960
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txt_fc 
         Height          =   360
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label50 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   600
         TabIndex        =   87
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label49 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   600
         TabIndex        =   86
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label48 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   600
         TabIndex        =   85
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "MPa"
         Height          =   255
         Left            =   3000
         TabIndex        =   39
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "MPa"
         Height          =   255
         Left            =   3000
         TabIndex        =   38
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "MPa"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "B1 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "fyt :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "f y  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "f 'c : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Label Label57 
      BackStyle       =   0  'Transparent
      Caption         =   "* indicates a required field."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9480
      TabIndex        =   97
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Singly Reinforce Beam Design"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   13695
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_calculate_Click()
    If Not (txt_mainbar.Text = "") And Not (txt_stirrups.Text = "") And Not (txt_height.Text = "") And Not (txt_fc.Text = "") And Not (txt_musup.Text = "") And Not (txt_mumid.Text = "") And Not (txt_shear.Text = "") And Not (txt_fy.Text = "") And Not (txt_fyt.Text = "") Then
        computeB1
        computeRebars
        computeDepth
        computeBeamAdequacy
        computeAtSupport
        computeAtMidspan
        computeStirrups
        btn_details.Enabled = True
        btn_viewStirrups.Enabled = True
    Else
        MsgBox "Please enter some value first. (f'c, fy, fyt, Main Bar, Stirrups, base, height, MUsup MUmid and Shear). * indicates a required field.", vbInformation, "Warning"
    End If
    
End Sub

Private Sub btn_clearFields_Click()
    clearFields
End Sub

Private Sub btn_details_Click()
    Details_Form.Show
    computeAtSupport
    computeAtMidspan
    Details_Form.getNumberOfBarsAtSupport (Details_Form.txt_atSnumBars.Text)
    Details_Form.getNumberOfBarsAtMidspan (Details_Form.txt_atMnumBars.Text)
    
End Sub

Private Sub btn_viewStirrups_Click()
    Stirrups_Form.Show
    computeStirrups
End Sub


Function computeB1()
On Error GoTo ErrorHandler
    If Not (txt_fc.Text = "") Then
        If txt_fc.Text <= 28 Then
            txt_b1.Text = 0.85
        ElseIf (txt_fc.Text < 55) Then
            txt_b1.Text = Format((0.85 - (0.05 / 7)) * (txt_fc.Text - 28), "0.00")
        Else
            txt_b1.Text = 0.65
        End If
    Else
        MsgBox "Please enter value of f'c.", vbInformation, "Warning"
    End If
    Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"
End Function


Function computeRebars()
On Error GoTo ErrorHandler
    If Not (txt_mainbar.Text = "") And Not (txt_stirrups.Text = "") Then
        txt_nominalAreaAn.Text = Format((3.141592654 * (txt_mainbar.Text ^ 2)) / 4, "0")
    
        txt_nominalAreaAnt.Text = Format((3.141592654 * (txt_stirrups.Text ^ 2)) / 4, "0")
        
        If txt_mainbar.Text <= 16 Then
            txt_concreteCover.Text = 40
        Else
            txt_concreteCover.Text = 50
        End If
    Else
        MsgBox "Please enter value of Main Bars, db and Nominal Area, An.", vbInformation, "Warning"
    End If
    Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"
End Function

Function computeDepth()
On Error GoTo ErrorHandler
    If Not (txt_mainbar.Text = "") And Not (txt_stirrups.Text = "") And Not (txt_height.Text = "") Then
        txt_effectiveDepth.Text = txt_height.Text - txt_concreteCover.Text - txt_stirrups.Text - (txt_mainbar.Text / 2)
    Else
        MsgBox "Please enter value of Height,h and compute first Rebars Geometry.", vbInformation, "Warning"
    End If
    Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"
End Function

Function computeBeamAdequacy()
On Error GoTo ErrorHandler
    txt_pmax.Text = Format(((0.85 * Val(txt_fc.Text) * 0.85) / Val(txt_fy.Text)) * ((0.003) / (0.008)), "0.000")
    txt_pbal.Text = Format(((0.85 * Val(txt_fc.Text) * 0.85) / Val(txt_fy.Text)) * ((600) / (600 + Val(txt_fy.Text))), "0.000")
    txt_w.Text = Format(((Val(txt_pmax.Text)) * Val((Val(txt_fy.Text)) / (Val(txt_fc.Text)))), "0.000")
    txt_rn.Text = Format(((Val(txt_fc.Text)) * (Val(txt_w.Text))) * ((1 - (0.59 * Val(txt_w.Text)))), "0.000")
    txt_dreq.Text = Format(((Val(txt_musup.Text) * 1000000) / (0.9 * Val(txt_rn.Text) * Val(txt_base.Text))) ^ (0.5), "0.00")
    txt_hreq.Text = Format(Val(txt_concreteCover.Text) + Val(txt_stirrups.Text) + Val(txt_dreq.Text) + (Val(txt_mainbar.Text) / 2), "0.00")
    If txt_hreq.Text < txt_height.Text Then
        lbl_adeq_ok.Visible = True
        lbl_adeq_ok.Caption = "OKAY"
        lbl_adeq_ok.ForeColor = &H8000&
    Else
        lbl_adeq_ok.Visible = True
        lbl_adeq_ok.Caption = "Please adjust beam dimension!"
        lbl_adeq_ok.ForeColor = &HFF&
    End If
    Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"
End Function


Function computeStirrups()
On Error GoTo ErrorHandler
    Dim vn As Double
    Dim Av As Double
    Dim StrpsEffDepth As Double
    Dim Vc As Double
    Dim spacingStrps As Double
    Dim Vs As Double
    Dim sMaxComp As Double
    Dim sMax As Double
    Dim checkVs As Double
    
    vn = txt_shear.Text / 0.75
    Av = 2 * (3.141592654 / (4)) * ((txt_stirrups.Text) ^ 2)
    StrpsEffDepth = txt_height.Text - txt_concreteCover.Text - (txt_stirrups.Text / 2)
    Vc = (0.17 * (txt_fc.Text ^ 0.5) * (txt_base.Text) * (StrpsEffDepth)) / (1000)
    If vn > Vc Then
        Vs = vn - Vc
    Else
        Vs = 0
    End If
    
    If vn > Vc Then
        spacingStrps = (Av * txt_fyt.Text * StrpsEffDepth) / (Vs * 1000)
        Stirrups_Form.Label_vnvc.Caption = "Vn > Vc"
        Stirrups_Form.Label_usevnvc.Caption = "Therefore, need stirrups!"
    Else
        spacingStrps = (3 * Av * txt_fyt.Text) / (txt_base.Text * 1000)
        Stirrups_Form.Label_vnvc.Caption = "Vn < Vc"
        tirrups_Form.Label_usevnvc.Caption = "Therefore, no need for stirrups!"
    End If
    
    sMaxComp = ((txt_fc.Text ^ 0.5) * txt_base.Text * StrpsEffDepth) / (3 * 1000)
    
    If Vs < sMaxComp Then
        If ((StrpsEffDepth / 2) < 600) Then
            sMax = (StrpsEffDepth / 2)
        ElseIf (600 < (StrpsEffDepth / 2)) Then
            sMax = 600
        End If
    Else
        If ((StrpsEffDepth / 4) < 400) Then
            sMax = (StrpsEffDepth / 4)
        ElseIf (400 < (StrpsEffDepth / 4)) Then
            sMax = 400
        End If
    End If
    
    If sMax < spacingStrps Then
        Stirrups_Form.Label_smax.Caption = "Smax < S"
    Else
        Stirrups_Form.Label_smax.Caption = "Smax > S"
    End If
    
    If (sMax < spacingStrps) Then
        Stirrups_Form.Label_therefore.Caption = "Therefore, use S = " & sMax & " mm"
        txt_wr2.Text = "@ " & sMax & " mm O.C."
    ElseIf (spacingStrps < sMax) Then
        Stirrups_Form.Label_therefore.Caption = "Therefore, use S = " & spacingStrps & " mm"
        txt_wr2.Text = "@ " & spacingStrps & " mm O.C."
    End If
    
    checkVs = ((0.66) * (txt_fc.Text ^ 0.5) * txt_base.Text * StrpsEffDepth) / (1000)
    
    If (Vs < checkVs) Then
        Stirrups_Form.Label_check.Caption = "Beam section is economical."
        Stirrups_Form.Label_checklessthan = "Vs<="
    ElseIf (checkVs < Vs) Then
        Stirrups_Form.Label_check.Caption = "Please adjust Beam Section."
        Stirrups_Form.Label_checklessthan = "Vs>="
    End If
    
    Stirrups_Form.txt_vn.Text = Format(vn, "0.00")
    Stirrups_Form.txt_av.Text = Format(Av, "0.00")
    Stirrups_Form.txt_effdepth.Text = Format(StrpsEffDepth, "0")
    Stirrups_Form.txt_vc.Text = Format(Vc, "0.00")
    Stirrups_Form.txt_vs.Text = Format(Vs, "0.00")
    Stirrups_Form.txt_spacing.Text = Format(spacingStrps, "0.00")
    Stirrups_Form.txt_smax.Text = Format(sMax, "0.00")
    Stirrups_Form.txt_vsCheck.Text = Format(checkVs, "0.00")
    
    Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"
End Function

Public Function computeAtSupport()
On Error GoTo ErrorHandler
    Dim data0 As Double
    Dim compData1 As Double
    Dim compData2 As Double
    Dim minValueA As Double
    Dim c As Double
    Dim fs As Double
    Dim et As Double
    Dim diameter As Double
    Dim steelArea As Double
    
    Dim numBars As Double
    Dim clearSpacing As Double
    
    
    data0 = (0.9) * txt_b1.Text * txt_fc.Text * txt_base.Text
    compData1 = ((data0 * Val(txt_effectiveDepth.Text)) + ((((data0 * Val(txt_effectiveDepth.Text)) ^ 2) - (4 * (data0 / 2) * (Val(txt_musup.Text) * 1000000))) ^ (1 / 2))) / data0
    compData2 = ((data0 * Val(txt_effectiveDepth.Text)) - ((((data0 * Val(txt_effectiveDepth.Text)) ^ 2) - (4 * (data0 / 2) * (Val(txt_musup.Text) * 1000000))) ^ (1 / 2))) / data0
    
    If (compData1 < compData2) Then
        minValueA = compData1
    ElseIf (compData2 < compData1) Then
        minValueA = compData2
    End If
    
    c = minValueA / txt_b1.Text
    
    fs = 600 * ((txt_effectiveDepth.Text - c) / c)
    
    et = Format(0.003 * ((txt_effectiveDepth.Text - c) / c), "0.0000")
    
    If (et < 0.005) Then
        diameter = 0.65
    Else
        diameter = 0.9
    End If
    
    If (fs > txt_fy.Text) Then
        steelArea = ((Val(txt_musup.Text * 1000000)) / ((diameter) * (txt_fy.Text) * ((Val(txt_effectiveDepth.Text)) - (minValueA / 2))))
    Else
        steelArea = ((Val(txt_musup.Text * 1000000)) / ((diameter) * (fs) * ((Val(txt_effectiveDepth.Text)) - (minValueA / 2))))
    End If
    
    numBars = RoundUp(((steelArea) / (txt_nominalAreaAn.Text)))
    
    If numBars = 1 Then
        numBars = 2
    End If
    
    clearSpacing = (((txt_base.Text) - (2 * txt_concreteCover.Text) - (2 * txt_stirrups.Text) - (txt_mainbar.Text)) / (numBars - 1)) - (txt_mainbar.Text)
    txt_srAtSupport.Text = numBars & " - " & txt_mainbar.Text & " mm ÿ bar"
    txt_wr1.Text = txt_stirrups.Text & " mm ÿ bar"
    
    Details_Form.txt_atSa.Text = Format(minValueA, "0.00")
    Details_Form.txt_atSc.Text = Format(c, "0.00")
    Details_Form.txt_atSfs.Text = Format(fs, "0")
    If fs > txt_fy.Text Then
        Details_Form.txt_atSfsfy.Caption = "fs > fy"
        Details_Form.txt_atSusefsfy.Caption = "Use fy!"
    Else
        Details_Form.txt_atSfsfy.Caption = "fs < fy"
        Details_Form.txt_atSusefsfy.Caption = "Use fs!"
    End If
    
    Details_Form.txt_atSEt.Text = Format(et, "0.0000")
    If et < 0.005 Then
        Details_Form.txt_atS005.Caption = "< 0.005"
        Details_Form.txt_atSctControlled.Caption = "Therefore, it is Compression Controlled!"
        Details_Form.txt_atSuseDia.Caption = "Use ÿ = " & Format(diameter, "0.00")
    Else
        Details_Form.txt_atS005.Caption = "> 0.005"
        Details_Form.txt_atSctControlled.Caption = "Therefore, it is Tension Controlled!"
        Details_Form.txt_atSuseDia.Caption = "Use ÿ = " & Format(diameter, "0.00")
    End If
    
    Details_Form.txt_atSsteelArea.Text = Format(steelArea, "0")
    Details_Form.txt_atSnumBars.Text = Format(numBars, "0")
    Details_Form.txt_atSbars.Caption = txt_mainbar.Text & " mm ÿ bar"
    
    Details_Form.txt_atSclearSpacing.Text = Format(clearSpacing, "0")
    
    Details_Form.Label_Sbase.Caption = txt_base.Text & " mm"
    Details_Form.Label_Sheight.Caption = txt_height.Text & " mm"
    Details_Form.Label_Sdepth.Caption = txt_effectiveDepth.Text & " mm"
    
    
    Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"
    
End Function

Function computeAtMidspan()
On Error GoTo ErrorHandler
    Dim data0 As Double
    Dim compData1 As Double
    Dim compData2 As Double
    Dim minValueA As Double
    Dim c As Double
    Dim fs As Double
    Dim et As Double
    Dim diameter As Double
    Dim steelArea As Double
    
    Dim numBars As Double
    Dim clearSpacing As Double
    
    
    data0 = (0.9) * txt_b1.Text * txt_fc.Text * txt_base.Text
    compData1 = ((data0 * Val(txt_effectiveDepth.Text)) + ((((data0 * Val(txt_effectiveDepth.Text)) ^ 2) - (4 * (data0 / 2) * (Val(txt_mumid.Text) * 1000000))) ^ (1 / 2))) / data0
    compData2 = ((data0 * Val(txt_effectiveDepth.Text)) - ((((data0 * Val(txt_effectiveDepth.Text)) ^ 2) - (4 * (data0 / 2) * (Val(txt_mumid.Text) * 1000000))) ^ (1 / 2))) / data0
    
    If (compData1 < compData2) Then
        minValueA = compData1
    ElseIf (compData2 < compData1) Then
        minValueA = compData2
    End If
    
    c = minValueA / txt_b1.Text
    
    fs = 600 * ((txt_effectiveDepth.Text - c) / c)
    
    et = Format(0.003 * ((txt_effectiveDepth.Text - c) / c), "0.0000")
    
    If (et < 0.005) Then
        diameter = 0.65
    Else
        diameter = 0.9
    End If
    
    If (fs > txt_fy.Text) Then
        steelArea = ((Val(txt_mumid.Text * 1000000)) / ((diameter) * (txt_fy.Text) * ((Val(txt_effectiveDepth.Text)) - (minValueA / 2))))
    Else
        steelArea = ((Val(txt_mumid.Text * 1000000)) / ((diameter) * (fs) * ((Val(txt_effectiveDepth.Text)) - (minValueA / 2))))
    End If
    
    numBars = RoundUp(((steelArea) / (txt_nominalAreaAn.Text)))
    
    If numBars = 1 Then
        numBars = 2
    End If
    
    clearSpacing = (((txt_base.Text) - (2 * txt_concreteCover.Text) - (2 * txt_stirrups.Text) - (txt_mainbar.Text)) / (numBars - 1)) - (txt_mainbar.Text)
    txt_srAtMidspan.Text = numBars & " - " & txt_mainbar.Text & " mm ÿ bar"
    txt_wr1.Text = txt_stirrups.Text & " mm ÿ bar"
    
    Details_Form.txt_atMa.Text = Format(minValueA, "0.00")
    Details_Form.txt_atMc.Text = Format(c, "0.00")
    Details_Form.txt_atMfs.Text = Format(fs, "0")
    If fs > txt_fy.Text Then
        Details_Form.txt_atMfsfy.Caption = "fs > fy"
        Details_Form.txt_atMusefsfy.Caption = "Use fy!"
    Else
        Details_Form.txt_atMfsfy.Caption = "fs < fy"
        Details_Form.txt_atMusefsfy.Caption = "Use fs!"
    End If
    
    Details_Form.txt_atMEt.Text = Format(et, "0.0000")
    If et < 0.005 Then
        Details_Form.txt_atM005.Caption = "< 0.005"
        Details_Form.txt_atMctControlled.Caption = "Therefore, it is Compression Controlled!"
        Details_Form.txt_atMuseDia.Caption = "Use ÿ = " & Format(diameter, "0.00")
    Else
        Details_Form.txt_atM005.Caption = "> 0.005"
        Details_Form.txt_atMctControlled.Caption = "Therefore, it is Tension Controlled!"
        Details_Form.txt_atMuseDia.Caption = "Use ÿ = " & Format(diameter, "0.00")
    End If
    
    Details_Form.txt_atMsteelArea.Text = Format(steelArea, "0")
    Details_Form.txt_atMnumBars.Text = Format(numBars, "0")
    Details_Form.txt_atMbars.Caption = txt_mainbar.Text & " mm ÿ bar"
    
    Details_Form.txt_atMclearSpacing.Text = Format(clearSpacing, "0")
    
    Details_Form.Label_Mbase.Caption = txt_base.Text & " mm"
    Details_Form.Label_Mheight.Caption = txt_height.Text & " mm"
    Details_Form.Label_Mdepth.Caption = txt_effectiveDepth.Text & " mm"
    
    Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"
End Function

Public Function RoundUp(ByVal Value As Double) As Double
Dim temp As Double
    temp = Int(Value)
    If temp <> Value Then
        temp = temp + 1
    End If
    RoundUp = temp
End Function


Function clearFields()
    txt_fc.Text = ""
    txt_fy.Text = ""
    txt_fyt.Text = ""
    txt_b1.Text = 0
    
    txt_mainbar.Text = ""
    txt_nominalAreaAn = 0
    txt_stirrups.Text = ""
    txt_nominalAreaAnt.Text = 0
    txt_concreteCover.Text = 0
    
    txt_base.Text = ""
    txt_height.Text = ""
    txt_length.Text = ""
    txt_effectiveDepth.Text = 0
    
    txt_musup.Text = ""
    txt_mumid.Text = ""
    txt_shear.Text = ""
    
    txt_pmax.Text = 0
    txt_pbal.Text = 0
    txt_w.Text = 0
    txt_rn.Text = 0
    txt_dreq.Text = 0
    txt_hreq.Text = 0
    lbl_adeq_ok.Visible = False
    
    txt_srAtSupport.Text = 0
    txt_srAtMidspan.Text = 0
    txt_wr1.Text = 0
    txt_wr2.Text = 0
    btn_details.Enabled = False
    btn_viewStirrups.Enabled = False
    
End Function



Private Sub Command1_Click()
    computeB1
End Sub

Private Sub Command2_Click()
    computeRebars
End Sub

Private Sub Command3_Click()
    computeDepth
End Sub

Private Sub Command4_Click()
    computeAtMidspan
End Sub
