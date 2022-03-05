VERSION 5.00
Begin VB.Form Details_Form 
   Caption         =   "Steel Reinforcements Details"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Calculations : "
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
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   12255
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   " At Midspan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   5775
         Begin VB.TextBox txt_atMEt 
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
            Left            =   1200
            TabIndex        =   51
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txt_atMfs 
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
            Left            =   1200
            TabIndex        =   50
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txt_atMc 
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
            Left            =   1200
            TabIndex        =   49
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txt_atMa 
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
            Left            =   1200
            TabIndex        =   48
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txt_atMclearSpacing 
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
            Left            =   2400
            TabIndex        =   43
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox txt_atMnumBars 
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
            Left            =   2400
            TabIndex        =   40
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txt_atMsteelArea 
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
            Left            =   2400
            TabIndex        =   28
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label txt_atM005 
            BackStyle       =   0  'Transparent
            Caption         =   "0.005"
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
            Left            =   2640
            TabIndex        =   55
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "MPa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   54
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   53
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   52
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   45
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Clear Spacing, Cs : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   44
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label txt_atMbars 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   42
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Bars : "
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
            Left            =   600
            TabIndex        =   41
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "a : "
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
            Left            =   600
            TabIndex        =   39
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "c : "
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
            Left            =   600
            TabIndex        =   38
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Checking : "
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
            Left            =   600
            TabIndex        =   37
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "fs : "
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
            Left            =   600
            TabIndex        =   36
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label txt_atMfsfy 
            BackStyle       =   0  'Transparent
            Caption         =   "fsfy : "
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
            Left            =   3240
            TabIndex        =   35
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label txt_atMusefsfy 
            BackStyle       =   0  'Transparent
            Caption         =   "use fsfy : "
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
            Left            =   3240
            TabIndex        =   34
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Et : "
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
            Left            =   600
            TabIndex        =   33
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label txt_atMctControlled 
            BackStyle       =   0  'Transparent
            Caption         =   "CTControlled : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   32
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label txt_atMuseDia 
            BackStyle       =   0  'Transparent
            Caption         =   "use dia : "
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
            Left            =   3240
            TabIndex        =   31
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "mmsq."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   30
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Steel Area, As : "
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
            Left            =   600
            TabIndex        =   29
            Top             =   2760
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   " At Support : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   5775
         Begin VB.TextBox txt_atSclearSpacing 
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
            Left            =   2280
            TabIndex        =   25
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox txt_atSnumBars 
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
            Left            =   2280
            TabIndex        =   22
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txt_atSsteelArea 
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
            Left            =   2280
            TabIndex        =   19
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox txt_atSEt 
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
            Left            =   1080
            TabIndex        =   15
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txt_atSfs 
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
            Left            =   1080
            TabIndex        =   11
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txt_atSc 
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
            Left            =   1080
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txt_atSa 
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
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "MPa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   47
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label txt_atS005 
            BackStyle       =   0  'Transparent
            Caption         =   "0.005"
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
            Left            =   2520
            TabIndex        =   46
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   27
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Clear Spacing, Cs : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label txt_atSbars 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   24
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Bars : "
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
            Left            =   480
            TabIndex        =   23
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Steel Area, As : "
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
            Left            =   480
            TabIndex        =   21
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "mmsq."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   20
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label txt_atSuseDia 
            BackStyle       =   0  'Transparent
            Caption         =   "use dia : "
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
            Left            =   3120
            TabIndex        =   18
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label txt_atSctControlled 
            BackStyle       =   0  'Transparent
            Caption         =   "CTControlled : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3120
            TabIndex        =   17
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Et : "
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
            Left            =   480
            TabIndex        =   16
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label txt_atSusefsfy 
            BackStyle       =   0  'Transparent
            Caption         =   "use fsfy : "
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
            Left            =   3120
            TabIndex        =   14
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label txt_atSfsfy 
            BackStyle       =   0  'Transparent
            Caption         =   "fsfy : "
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
            Left            =   3120
            TabIndex        =   13
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "fs : "
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
            Left            =   480
            TabIndex        =   12
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Checking : "
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
            Left            =   480
            TabIndex        =   10
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "c : "
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
            Left            =   480
            TabIndex        =   9
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   8
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "a : "
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
            Left            =   480
            TabIndex        =   6
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   5
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.Label Label_Mdepth 
      Caption         =   "Depth"
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
      Left            =   11640
      TabIndex        =   61
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label_Sdepth 
      Caption         =   "Depth"
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
      Left            =   5400
      TabIndex        =   60
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Shape Shape10 
      BorderWidth     =   2
      Height          =   2535
      Left            =   11520
      Top             =   1200
      Width           =   15
   End
   Begin VB.Line Line24 
      BorderWidth     =   3
      X1              =   11520
      X2              =   11400
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line23 
      BorderWidth     =   3
      X1              =   11520
      X2              =   11640
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line22 
      BorderWidth     =   3
      X1              =   11640
      X2              =   11520
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line21 
      BorderWidth     =   3
      X1              =   11400
      X2              =   11520
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Shape Shape9 
      BorderWidth     =   2
      Height          =   2535
      Left            =   5280
      Top             =   1200
      Width           =   15
   End
   Begin VB.Line Line20 
      BorderWidth     =   3
      X1              =   5280
      X2              =   5160
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line19 
      BorderWidth     =   3
      X1              =   5280
      X2              =   5400
      Y1              =   3720
      Y2              =   3600
   End
   Begin VB.Line Line18 
      BorderWidth     =   3
      X1              =   5400
      X2              =   5280
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line17 
      BorderWidth     =   3
      X1              =   5160
      X2              =   5280
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Label Label_Mheight 
      Alignment       =   1  'Right Justify
      Caption         =   "Height"
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
      Left            =   6240
      TabIndex        =   59
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label_Sheight 
      Alignment       =   1  'Right Justify
      Caption         =   "Height"
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
      Left            =   0
      TabIndex        =   58
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label_Mbase 
      Alignment       =   2  'Center
      Caption         =   "Base"
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
      Left            =   7920
      TabIndex        =   57
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label_Sbase 
      Alignment       =   2  'Center
      Caption         =   "Base"
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
      Left            =   1680
      TabIndex        =   56
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Shape Shape8 
      BorderWidth     =   2
      Height          =   2895
      Left            =   7320
      Top             =   1200
      Width           =   15
   End
   Begin VB.Line Line16 
      BorderWidth     =   3
      X1              =   7320
      X2              =   7200
      Y1              =   4080
      Y2              =   3960
   End
   Begin VB.Line Line15 
      BorderWidth     =   3
      X1              =   7320
      X2              =   7440
      Y1              =   4080
      Y2              =   3960
   End
   Begin VB.Line Line14 
      BorderWidth     =   3
      X1              =   7440
      X2              =   7320
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      X1              =   7200
      X2              =   7320
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   2
      Height          =   2895
      Left            =   1080
      Top             =   1200
      Width           =   15
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   1080
      X2              =   960
      Y1              =   4080
      Y2              =   3960
   End
   Begin VB.Line Line11 
      BorderWidth     =   3
      X1              =   1080
      X2              =   1200
      Y1              =   4080
      Y2              =   3960
   End
   Begin VB.Line Line10 
      BorderWidth     =   3
      X1              =   1200
      X2              =   1080
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   960
      X2              =   1080
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   7560
      X2              =   7680
      Y1              =   4200
      Y2              =   4080
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   7680
      X2              =   7560
      Y1              =   4320
      Y2              =   4200
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      X1              =   11280
      X2              =   11160
      Y1              =   4200
      Y2              =   4080
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   11160
      X2              =   11280
      Y1              =   4320
      Y2              =   4200
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   2
      Height          =   15
      Left            =   7560
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   1320
      X2              =   1440
      Y1              =   4200
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   1440
      X2              =   1320
      Y1              =   4320
      Y2              =   4200
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   5040
      X2              =   4920
      Y1              =   4200
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   4920
      X2              =   5040
      Y1              =   4320
      Y2              =   4200
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   15
      Left            =   1320
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   2895
      Left            =   7560
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Shape Shape6 
      BorderWidth     =   2
      Height          =   2055
      Left            =   8040
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2895
      Left            =   1320
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   2055
      Left            =   1800
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Steel Reinforcements Details"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   12495
   End
End
Attribute VB_Name = "Details_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function getNumberOfBarsAtSupport(numBars As Integer)
On Error GoTo ErrorHandler
    
    FillStyle = vbSolid
    FillColor = vbBlack
    DrawWidth = 3
    
    Select Case numBars
        Case 1: Circle (1920, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 2: Circle (1920, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 3: Circle (1920, 3590), 100, vbBlack
                Circle (3200, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 4: Circle (1920, 3590), 100, vbBlack
                Circle (2780, 3590), 100, vbBlack
                Circle (3630, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 5: Circle (1920, 3590), 100, vbBlack
                Circle (2555, 3590), 100, vbBlack
                Circle (3180, 3590), 100, vbBlack
                Circle (3810, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 6: Circle (1920, 3590), 100, vbBlack
                Circle (2395, 3590), 100, vbBlack
                Circle (2910, 3590), 100, vbBlack
                Circle (3425, 3590), 100, vbBlack
                Circle (3940, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 7: Circle (1920, 3590), 100, vbBlack
                Circle (2350, 3590), 100, vbBlack
                Circle (2765, 3590), 100, vbBlack
                Circle (3180, 3590), 100, vbBlack
                Circle (3600, 3590), 100, vbBlack
                Circle (4025, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 8: Circle (1920, 3590), 100, vbBlack
                Circle (2290, 3590), 100, vbBlack
                Circle (2645, 3590), 100, vbBlack
                Circle (3010, 3590), 100, vbBlack
                Circle (3370, 3590), 100, vbBlack
                Circle (3735, 3590), 100, vbBlack
                Circle (4100, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 9: Circle (1920, 3590), 100, vbBlack
                Circle (2250, 3590), 100, vbBlack
                Circle (2575, 3590), 100, vbBlack
                Circle (2890, 3590), 100, vbBlack
                Circle (3200, 3590), 100, vbBlack
                Circle (3515, 3590), 100, vbBlack
                Circle (3830, 3590), 100, vbBlack
                Circle (4130, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case 10: Circle (1920, 3590), 100, vbBlack
                Circle (2207, 3590), 100, vbBlack
                Circle (2480, 3590), 100, vbBlack
                Circle (2767, 3590), 100, vbBlack
                Circle (3045, 3590), 100, vbBlack
                Circle (3325, 3590), 100, vbBlack
                Circle (3605, 3590), 100, vbBlack
                Circle (3880, 3590), 100, vbBlack
                Circle (4155, 3590), 100, vbBlack
                Circle (4430, 3590), 100, vbBlack
        Case Else: MsgBox "Please select another section."
   End Select
   Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"

End Function


Public Function getNumberOfBarsAtMidspan(numBars As Integer)
On Error GoTo ErrorHandler
    
    FillStyle = vbSolid
    FillColor = vbBlack
    DrawWidth = 3
    
    Select Case numBars
        Case 1: Circle (8175, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 2: Circle (8175, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 3: Circle (8175, 3590), 100, vbBlack
                Circle (9455, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 4: Circle (8175, 3590), 100, vbBlack
                Circle (9035, 3590), 100, vbBlack
                Circle (9885, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 5: Circle (8175, 3590), 100, vbBlack
                Circle (8810, 3590), 100, vbBlack
                Circle (9435, 3590), 100, vbBlack
                Circle (10065, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 6: Circle (8175, 3590), 100, vbBlack
                Circle (8650, 3590), 100, vbBlack
                Circle (9165, 3590), 100, vbBlack
                Circle (9680, 3590), 100, vbBlack
                Circle (10195, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 7: Circle (8175, 3590), 100, vbBlack
                Circle (8605, 3590), 100, vbBlack
                Circle (9020, 3590), 100, vbBlack
                Circle (9435, 3590), 100, vbBlack
                Circle (9855, 3590), 100, vbBlack
                Circle (10280, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 8: Circle (8175, 3590), 100, vbBlack
                Circle (8545, 3590), 100, vbBlack
                Circle (8900, 3590), 100, vbBlack
                Circle (9265, 3590), 100, vbBlack
                Circle (9625, 3590), 100, vbBlack
                Circle (9990, 3590), 100, vbBlack
                Circle (10355, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 9: Circle (8175, 3590), 100, vbBlack
                Circle (8505, 3590), 100, vbBlack
                Circle (8830, 3590), 100, vbBlack
                Circle (9145, 3590), 100, vbBlack
                Circle (9455, 3590), 100, vbBlack
                Circle (9770, 3590), 100, vbBlack
                Circle (10085, 3590), 100, vbBlack
                Circle (10385, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case 10: Circle (8175, 3590), 100, vbBlack
                Circle (8462, 3590), 100, vbBlack
                Circle (8735, 3590), 100, vbBlack
                Circle (9022, 3590), 100, vbBlack
                Circle (9300, 3590), 100, vbBlack
                Circle (9580, 3590), 100, vbBlack
                Circle (9860, 3590), 100, vbBlack
                Circle (10135, 3590), 100, vbBlack
                Circle (10410, 3590), 100, vbBlack
                Circle (10680, 3590), 100, vbBlack
        Case Else: MsgBox "Please select another section."
   End Select
   Exit Function
ErrorHandler:
        MsgBox Err.Description, , "Error"

End Function

