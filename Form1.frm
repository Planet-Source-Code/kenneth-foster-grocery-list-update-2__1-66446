VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHiddenData 
      Height          =   435
      Left            =   60
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   53
      Top             =   5025
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picGreenBar 
      Height          =   390
      Left            =   600
      ScaleHeight     =   330
      ScaleWidth      =   555
      TabIndex        =   52
      Top             =   5025
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame frmEdit 
      BackColor       =   &H00F6F0E0&
      Caption         =   "Edit List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4680
      Left            =   5805
      TabIndex        =   24
      Top             =   165
      Visible         =   0   'False
      Width           =   6045
      Begin VB.CheckBox chkTT 
         BackColor       =   &H80000007&
         Caption         =   "Don't show Tooltips"
         Height          =   195
         Left            =   4005
         TabIndex        =   55
         Top             =   3960
         Width           =   195
      End
      Begin VB.TextBox txtInst 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Height          =   2370
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   615
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.TextBox txtEditDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F0E0&
         Height          =   3780
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   390
         Width           =   1800
      End
      Begin VB.TextBox txtEditCat 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F0E0&
         Height          =   3330
         Left            =   1890
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   390
         Width           =   2085
      End
      Begin VB.TextBox txtEditItems 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F0E0&
         Height          =   3330
         Left            =   3990
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   390
         Width           =   1950
      End
      Begin VB.TextBox txtNewCatName 
         Appearance      =   0  'Flat
         BackColor       =   &H00F6F0E0&
         Height          =   285
         Left            =   1935
         TabIndex        =   32
         Top             =   3945
         Width           =   1710
      End
      Begin Project1.MorphButton cmdEditInstr 
         Height          =   330
         Left            =   4980
         TabIndex        =   26
         Top             =   4275
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   582
         Caption         =   "Instructions"
         BackColor2      =   33023
         CaptionColor    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignment=   1
         CaptionHover    =   12632256
      End
      Begin Project1.MorphButton cmdSaveItems 
         Height          =   330
         Left            =   3975
         TabIndex        =   27
         Top             =   4275
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   582
         Caption         =   "Save"
         BackColor2      =   33023
         CaptionColor    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignment=   1
         CaptionHover    =   12632256
      End
      Begin Project1.MorphButton cmdEditSaveCat 
         Height          =   330
         Left            =   2820
         TabIndex        =   28
         Top             =   4275
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         Caption         =   "Save"
         BackColor2      =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignment=   1
         CaptionHover    =   8421504
      End
      Begin Project1.MorphButton cmdEditDelCat 
         Height          =   330
         Left            =   1905
         TabIndex        =   29
         Top             =   4275
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   582
         Caption         =   "Delete"
         BackColor2      =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignment=   1
         CaptionHover    =   8421504
      End
      Begin Project1.MorphButton cmdEditSaveDesc 
         Height          =   330
         Left            =   870
         TabIndex        =   30
         Top             =   4275
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         Caption         =   "Save"
         BackColor2      =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignment=   1
         CaptionHover    =   8421504
      End
      Begin Project1.MorphButton cmdEditLoadDesc 
         Height          =   330
         Left            =   75
         TabIndex        =   31
         Top             =   4275
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         Caption         =   "Load"
         BackColor2      =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAlignment=   1
         CaptionHover    =   8421504
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Tooltips"
         Height          =   210
         Left            =   4335
         TabIndex        =   56
         Top             =   3960
         Width           =   1485
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "New Category Name"
         Height          =   210
         Left            =   2055
         TabIndex        =   39
         Top             =   3735
         Width           =   1515
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Left            =   45
         Top             =   4245
         Width           =   5940
      End
      Begin VB.Shape Shape4 
         Height          =   4680
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   6045
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Descriptors"
         Height          =   210
         Left            =   315
         TabIndex        =   38
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Categories"
         Height          =   240
         Left            =   2535
         TabIndex        =   37
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Items and Prices"
         Height          =   210
         Left            =   4275
         TabIndex        =   36
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkGrid 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Grid"
      Height          =   225
      Left            =   9000
      TabIndex        =   17
      Top             =   4575
      Width           =   225
   End
   Begin VB.ListBox lstRecipe 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F0E0&
      Height          =   1980
      Left            =   2385
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   5115
      Width           =   3150
   End
   Begin Project1.MorphButton cmdExit 
      Height          =   255
      Left            =   11055
      TabIndex        =   14
      Top             =   165
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   450
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
   End
   Begin VB.TextBox txtRecipeName 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F0E0&
      Height          =   285
      Left            =   135
      TabIndex        =   10
      Top             =   6000
      Width           =   1950
   End
   Begin VB.ComboBox cboUnits 
      BackColor       =   &H00F6F0E0&
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1830
   End
   Begin VB.ComboBox cboCat 
      BackColor       =   &H00F6F0E0&
      Height          =   315
      Left            =   195
      TabIndex        =   1
      Top             =   1620
      Width           =   1920
   End
   Begin VB.TextBox txtQuan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F0E0&
      Height          =   285
      Left            =   345
      TabIndex        =   0
      Text            =   "1"
      Top             =   435
      Width           =   360
   End
   Begin MSComctlLib.ListView ltbItem 
      Height          =   4440
      Left            =   2370
      TabIndex        =   8
      Top             =   405
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   7832
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16183520
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin RichTextLib.RichTextBox rtfmain 
      Height          =   2685
      Left            =   5895
      TabIndex        =   15
      Top             =   5130
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   4736
      _Version        =   393217
      BackColor       =   16183520
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":40FB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4080
      Left            =   5850
      TabIndex        =   18
      Top             =   435
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   7197
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin Project1.MorphButton cmdClearSel 
      Height          =   270
      Left            =   300
      TabIndex        =   40
      Top             =   2010
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   476
      Caption         =   "Clear Selection"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   8421504
      FocusEnabled    =   -1  'True
   End
   Begin Project1.MorphButton cmdPrintList 
      Height          =   390
      Left            =   525
      TabIndex        =   41
      Top             =   4410
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      Caption         =   "Print List"
      BackColor2      =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   8421504
      FocusEnabled    =   -1  'True
   End
   Begin Project1.MorphButton cmdEdit 
      Height          =   390
      Left            =   525
      TabIndex        =   42
      Top             =   3945
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      Caption         =   "Edit (Show)"
      BackColor1      =   8421504
      BackColor2      =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   8421504
      FocusEnabled    =   -1  'True
   End
   Begin Project1.MorphButton cmdClear 
      Height          =   390
      Left            =   525
      TabIndex        =   43
      Top             =   3480
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      Caption         =   "Clear List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   8421504
      FocusEnabled    =   -1  'True
   End
   Begin Project1.MorphButton cmdDeleteItem 
      Height          =   390
      Left            =   525
      TabIndex        =   44
      Top             =   3060
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      Caption         =   "Delete Item"
      CaptionColor    =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   8421631
      FocusEnabled    =   -1  'True
   End
   Begin Project1.MorphButton cmdAdd 
      Height          =   390
      Left            =   540
      TabIndex        =   45
      Top             =   2640
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      Caption         =   "Add Item"
      BackColor2      =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   8421504
      FocusEnabled    =   -1  'True
   End
   Begin Project1.MorphButton cmdPrintr 
      Height          =   375
      Left            =   4755
      TabIndex        =   46
      Top             =   7395
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      Caption         =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   12632256
   End
   Begin Project1.MorphButton cmdCancelr 
      Height          =   375
      Left            =   3780
      TabIndex        =   47
      Top             =   7395
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "Cancel"
      BackColor1      =   4210752
      BackColor2      =   16777215
      CaptionColor    =   12632064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   16777088
   End
   Begin Project1.MorphButton cmdEditr 
      Height          =   375
      Left            =   2865
      TabIndex        =   48
      Top             =   7395
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "Edit"
      BackColor1      =   0
      BackColor2      =   16777088
      CaptionColor    =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   12648447
   End
   Begin Project1.MorphButton cmdDeleter 
      Height          =   375
      Left            =   1950
      TabIndex        =   49
      Top             =   7395
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "Delete"
      CaptionColor    =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   8421631
   End
   Begin Project1.MorphButton cmdSaver 
      Height          =   375
      Left            =   1035
      TabIndex        =   50
      Top             =   7395
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "Save"
      CaptionColor    =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   16744703
   End
   Begin Project1.MorphButton cmdAddr 
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   7395
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAlignment=   1
      CaptionHover    =   12632256
   End
   Begin Project1.DynamicPopupMenu DPM1 
      Left            =   75
      Top             =   5295
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin VB.Shape Shape1 
      Height          =   420
      Left            =   90
      Top             =   7365
      Width           =   5565
   End
   Begin VB.Shape Shape2 
      Height          =   2265
      Left            =   465
      Top             =   2580
      Width           =   1680
   End
   Begin VB.Label lblEditCatName 
      Height          =   240
      Left            =   1110
      TabIndex        =   54
      Top             =   5385
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   10875
      TabIndex        =   23
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Grocery List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8205
      TabIndex        =   22
      Top             =   180
      Width           =   1305
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10260
      TabIndex        =   21
      Top             =   4575
      Width           =   585
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9240
      TabIndex        =   20
      Top             =   4590
      Width           =   945
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5910
      TabIndex        =   19
      Top             =   4590
      Width           =   2865
   End
   Begin VB.Label lblTotRecords 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1410
      TabIndex        =   13
      Top             =   6465
      Width           =   645
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "New Recipe Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   285
      TabIndex        =   12
      Top             =   5775
      Width           =   1665
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Recipes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   150
      TabIndex        =   11
      Top             =   6495
      Width           =   1260
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Item List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3435
      TabIndex        =   9
      Top             =   150
      Width           =   900
   End
   Begin VB.Label lbltotUnitPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1365
      TabIndex        =   7
      Top             =   645
      Width           =   645
   End
   Begin VB.Image ImageUp 
      Height          =   240
      Left            =   750
      Picture         =   "Form1.frx":417D
      Top             =   360
      Width           =   240
   End
   Begin VB.Image ImageDown 
      Height          =   240
      Left            =   750
      Picture         =   "Form1.frx":42C7
      Top             =   585
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   375
      TabIndex        =   6
      Top             =   225
      Width           =   315
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Unit     Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1260
      TabIndex        =   5
      Top             =   210
      Width           =   870
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   765
      TabIndex        =   4
      Top             =   840
      Width           =   555
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Descriptors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   585
      TabIndex        =   3
      Top             =   1425
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'**                              My Grocery List
'**                               Version 5.1.1
'**                               By Ken Foster
'**                                Sept 4,2006
'**                     Freeware--- no copyrights claimed
'**       Special thanks to the authors of some of the code I used, whoever you are.
'*******************************************************************
'Sept 7, 2006
'Cleaned up code. Fixed a couple of minor bugs.
'Recipe counter and item edit window
'Added tooltips to Edit windows.
' Sept 5,2006
' Resized for 800 x 600 resolution
'=============================================

Option Explicit

 Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
 Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
   
   Dim TotPrice As Single
   Dim Price As String
   Dim firsttime As Boolean           'used to load RTF on start up
   Dim Rcount As Integer              'used to keep track of files that have (3) dots in front of them ex.  ...Filename. so recipe counter reflects correct number of recipes

Private Sub Form_Load()
Dim Region As Long
Dim strPath As String
Dim strMapName As String
Dim fStg As String
Dim fSLen As Integer
Dim firststg As String

Label16.Caption = Format(Now, "Long Date")         'set todays date
If Me.Picture <> 0 Then                                        'make form transparent
  Call SetAutoRgn(Me)
End If

'size the combo boxes in width and height
 MoveWindow cboUnits.hwnd, cboUnits.Left, cboUnits.Top, 110, 460, 1
 MoveWindow cboCat.hwnd, cboCat.Left, cboCat.Top, 130, 300, 1

'load recipes into listbox
lstRecipe.Clear
strPath = Dir(App.Path & "\RecipeFolder" & "\*.rtf")

If Not strPath = "" Then                                  'yes, there are files here so
   Do                                                             'go get them
      strMapName = strPath
      fSLen = Len(strMapName) - 4                    'filename length minus extension
      fStg = Mid$(strMapName, 1, fSLen)            'filename without extension
      lstRecipe.AddItem fStg                              'put filename into listbox

      firststg = Left$(fStg, 3)                              'don't count list item as a recipe if it starts with ...
      If firststg = "..." Then                                'in this case Instructions and Measurements
         Rcount = Rcount + 1
      End If
      lblTotRecords.Caption = lstRecipe.ListCount - Rcount           'Show how many recipes there are in list
      
      strPath = Dir$
   Loop Until strPath = ""
Else
   MsgBox "No files found!", vbCritical + vbOKOnly, "File - Error"
End If

   txtRecipeName.Locked = True                'disable textbox until needed
   ListViewSetup                                       ' columnheaders and greenbar background
   
   'load the lists
   LoadTextInstr
   LoadcboUnits
   LoadcboCat
   ListAll
   ' load Measurements file into RTB on startup
   firsttime = True
   RTB_Load
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Seld As String

If Button = 1 Then               'move form
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

If Button = 2 Then               'show popup menu
   Seld = DPM1.Popup("Delete Item,Clear List,Edit List,Print List,-,Exit")
   Select Case Seld
      Case "Delete Item":
          cmdDeleteItem_Click
      Case "Clear List":
          cmdClear_Click
      Case "Edit List":
          cmdEdit_Click
          If cmdEdit.Caption = "Edit (Hide)" Then
             cmdEdit.Caption = "Edit (Show)"
         Else
             cmdEdit.Caption = "Edit (Hide)"
         End If
      Case "Print List":
          cmdPrintList_Click
      Case "Exit":
         cmdExit_Click
   End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub chkGrid_Click()                             'show / hide grid
   If chkGrid.Value = Checked Then
      ListView1.GridLines = True
   Else
      ListView1.GridLines = False
   End If
End Sub

Private Sub cmdADD_Click()                            'add item
   Dim X As Integer
   
    If lbltotUnitPrice.Caption = "" Then Exit Sub
   'if no product name entered then clear everything and get out
   If ltbItem.SelectedItem.Text = "" Then
      txtQuan.Text = "1"
      cboUnits.Text = ""
      Exit Sub
   End If
   'quantity
    If txtQuan.Text = "" Then txtQuan.Text = "1"
    txtQuan.Text = Format(txtQuan.Text, "##0")
    X = ListView1.ListItems.Count + 1
    ListView1.ListItems.Add X, , txtQuan.Text
    ListView1.ListItems(X).SubItems(1) = cboUnits.Text
    ListView1.ListItems(X).SubItems(2) = ltbItem.SelectedItem.Text
    ListView1.ListItems(X).SubItems(3) = ltbItem.SelectedItem.SubItems(1)
    Price = Val(txtQuan.Text) * ltbItem.SelectedItem.SubItems(1)
    ListView1.ListItems(X).SubItems(4) = Format(Price, "$##0.00")
   
   TotPrice = TotPrice + Val(Price)                             ' update total
   Label1.Caption = Format(TotPrice, "$##0.00")
   
   'clear everything for next item
   ltbItem.SelectedItem.Selected = False
   txtQuan.Text = "1"
   txtQuan.SetFocus
   lbltotUnitPrice.Caption = ""
   ListView1.SelectedItem.Selected = False                 'hide the selection bar (highlight)
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50
End Sub

Private Sub cmdClear_Click()                                  'clear list
    ListView1.ListItems.Clear
    Label1.Caption = "$0.00"
    TotPrice = 0
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50
End Sub

Private Sub cmdClearSel_Click()                           'clears selections so another can be entered
   txtQuan.Text = "1"
   cboUnits.Text = ""
   lbltotUnitPrice.Caption = ""
   ltbItem.SelectedItem.Selected = False
   PcSpeakerBeep 500, 50
End Sub

Private Sub cmdDeleteItem_Click()                          'delete item
   On Error Resume Next
   If ListView1.SelectedItem.Selected = False Then Exit Sub
   If txtQuan.Text = "" Then txtQuan.Text = "1"
   If ListView1.ListItems.Count = 0 Then Exit Sub
   Price = Val(txtQuan.Text) * ListView1.SelectedItem.SubItems(4)
   ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
   TotPrice = TotPrice - Val(Price)
   Label1.Caption = Format(TotPrice, "$##0.00")
   PcSpeakerBeep 600, 50
   PcSpeakerBeep 400, 50
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
     frmEdit.Visible = Not frmEdit.Visible
     txtInst.Visible = False                                      'if instructions window was left open then close it
     txtEditCat.Text = ""
     txtEditDesc.Text = ""
     txtEditItems.Text = ""
     txtEditItems.Text = "Click on a category from the list on the left to show current items."
     cboCat_Click
     txtQuan.SetFocus
     EditLoadCat
     cmdEditLoadDesc_Click
     PcSpeakerBeep 400, 50
     PcSpeakerBeep 600, 50
     'keep border in position with edit frame
     Shape4.Top = 1
     Shape4.Left = 1
     Shape4.Visible = Not Shape4.Visible
End Sub

Private Sub cmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cmdEdit.Caption = "Edit (Hide)" Then
       cmdEdit.Caption = "Edit (Show)"
    Else
       cmdEdit.Caption = "Edit (Hide)"
    End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrintList_Click()                                  'print list
   Dim subit As String
   Dim i As Integer
    
    If ListView1.ListItems.Count = 0 Then Exit Sub
   On Error Resume Next
   Printer.ScaleMode = 3
   Printer.FontSize = 12
   Printer.Print Tab(25); "Grocery List"
   Printer.Print
   Printer.FontUnderline = True
   Printer.Print "Qty  Descriptor         Item                                                Price     Total Price"
   Printer.FontUnderline = False
    For i = 1 To ListView1.ListItems.Count
       subit = ListView1.ListItems(i).Text
       Printer.Print subit; Tab(6); ListView1.ListItems(i).SubItems(1); Tab(22); ListView1.ListItems(i).SubItems(2); Tab(59); ListView1.ListItems(i).SubItems(3); Tab(70); ListView1.ListItems(i).SubItems(4)
       subit = ""
    Next i
     Printer.Print
     Printer.Print Tab(62); "Total:  " & Label1.Caption
     Printer.NewPage
     Printer.EndDoc
     PcSpeakerBeep 400, 50
     PcSpeakerBeep 600, 50
   End Sub

Private Sub cmdADDr_Click()                                'clear and enable all needed controls
   'On Error Resume Next
    
   rtfmain.Text = ""
   rtfmain.Locked = False
   txtRecipeName.Locked = False
   cmdSaver.Enabled = True
   txtRecipeName.SetFocus
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50
End Sub

Private Sub cmdCancelr_Click()                             'clear and disable controls
   txtRecipeName.Locked = True
   txtRecipeName.Text = ""
   rtfmain.Text = ""
   rtfmain.Locked = True
   cmdSaver.Enabled = False
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50
End Sub

Private Sub cmdDeleter_Click()
  
   If rtfmain.Text = "" Then
      MsgBox "Nothing to Delete", vbInformation + vbOKOnly, "Delete Error"
      Exit Sub
   End If
   
  If MsgBox("Are you sure ?", vbInformation + vbYesNo, "Delete this file.") = vbYes Then
      Call File_Delete(lstRecipe)                                'delete the recipe from file
      Call List_Remove(lstRecipe)                                'delete the recipe from listbox
      lblTotRecords.Caption = lstRecipe.ListCount - Rcount
      rtfmain.Text = ""
      txtRecipeName.Locked = True
   End If
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50
End Sub

Private Sub cmdEditr_Click()
   If rtfmain.Text = "" Then
      MsgBox "Nothing to edit", vbInformation + vbOKOnly, "Edit Error"
      Exit Sub
   End If
   
   rtfmain.Locked = False
   txtRecipeName.Text = lstRecipe.Text
   cmdSaver.Enabled = True
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50
End Sub

Public Sub cmdPrintr_Click()  'this needs to be Public ,don't change or print button does'nt work on preview page
     
    ' Print the contents of the RichTextBox with a one inch margin
      On Error GoTo err1
      
      If rtfmain.Text = "" Then
         MsgBox "Nothing to Print", vbInformation + vbOKOnly, "No Recipe to Print"
         Exit Sub
      End If
      
      PrintRTFBox rtfmain, 1440, 1440, 1440, 1440                 '1440 Twips = 1 Inch
      Exit Sub
err1:
    Select Case Err.Number
        Case 482
            MsgBox "Make sure that you have a printer installed.  If a " & _
                "printer is installed, go into your printer properties " & _
                "look under the Setup tab, and make sure the ICM checkbox " & _
                "is checked and try printing again.", , "Printer Error"
            Exit Sub
        Case Else
            MsgBox Err.Number & " " & Err.Description
    End Select
    PcSpeakerBeep 400, 50
    PcSpeakerBeep 600, 50
End Sub

Private Sub cmdSaver_Click()
   
   If txtRecipeName.Text = "" Then
      MsgBox "Please enter a Name."
      Exit Sub
   End If
   
  ' File_Exists
   If FileExists(App.Path & "\RecipeFolder\" & txtRecipeName.Text & ".rtf") Then
      If MsgBox("File Exists!! Do you want to overwrite file?", vbYesNo, "File Exists") = vbNo Then Exit Sub
      Call RTB_Save                                                            'save updated recipe
   Else
      Call RTB_Save                                                            'save recipe
      lstRecipe.AddItem txtRecipeName.Text                         'add to listbox
      lblTotRecords.Caption = lstRecipe.ListCount - Rcount     'Show how many recipes there are in list
   End If
   
   'control logic
   txtRecipeName.Text = ""
   rtfmain.Text = ""
   rtfmain.Locked = True
   cmdSaver.Enabled = False
   txtRecipeName.Locked = True
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50
End Sub

Private Sub cmdEditInstr_Click()
   txtInst.Visible = Not txtInst.Visible
End Sub

Private Sub cmdEditLoadDesc_Click()
    LoadText (App.Path & "\ListData\UnitDes.udr"), txtEditDesc
End Sub

Private Sub cmdEditDelCat_Click()
    DeleteFile App.Path & "\ListData\" & txtEditCat.SelText & ".txt"
    txtEditItems.Text = ""
    txtNewCatName.Text = ""
    EditLoadCat
    cboCat.Clear
    LoadcboCat
End Sub

Private Sub cmdEditSaveCat_Click()
   If txtNewCatName.Text = "" Then
      MsgBox "Nothing to Save"
      Exit Sub
   End If
   SaveText "", App.Path & "\ListData\" & txtNewCatName.Text & ".txt"
   txtNewCatName.Text = ""
   txtEditCat.Text = ""
   cboCat.Clear
   EditLoadCat
   LoadcboCat
End Sub

Private Sub cmdEditSaveDesc_Click()
   SaveText txtEditDesc.Text, App.Path & "\ListData\UnitDes.udr"
   cboUnits.Clear
   LoadcboUnits
   txtEditDesc.Text = ""
End Sub

Private Sub cmdSaveItems_Click()
      If lblEditCatName.Caption = "" Then
         MsgBox "No category to save to"
         txtEditItems.Text = ""
         Exit Sub
      End If
      If MsgBox("Are you sure?", vbYesNo, "Save") = vbNo Then Exit Sub
    SaveText txtEditItems.Text, App.Path & "\ListData\" & lblEditCatName.Caption & ".txt"
    cboCat.Text = lblEditCatName.Caption
    cboCat_Click
End Sub

Private Sub cboCat_Click()
   If ListView1.ListItems.Count <> 0 Then ListView1.SelectedItem.Selected = False   'hide the selection bar (highlight)
   If cboCat.Text = "" Or cboCat.Text = "All" Then
      ListAll
   Else
      ltbItem.ListItems.Clear
      ltbItem.Sorted = False    'turn sorted off and load items into list
      LoadCombo cboCat.Text, cboCat
      ltbItem.Sorted = True     'with items loaded, we can now sort the items.Not all prices show if you don't do this
   End If
End Sub

Private Sub cboUnits_Change()
   If ListView1.ListItems.Count <> 0 Then ListView1.SelectedItem.Selected = False     'hide the selection bar (highlight)
End Sub

Private Sub ListAll()   'loads items from all catogories into one list
   Dim xxx As Integer
   'load all items in lstItemName
   ltbItem.ListItems.Clear
   ltbItem.Sorted = False
   For xxx = 1 To cboCat.ListCount - 1
      LoadCombo cboCat.List(xxx), cboCat
   Next xxx
   ltbItem.Sorted = True
   cboCat.Text = "All"
End Sub

Private Sub LoadCombo(textfile As String, cboname As ComboBox)  'loads items into listbox
   Dim strArray() As String
   Dim i As Integer
   Dim iFile As Integer
   Dim Y As Integer
   
      iFile = FreeFile
      If textfile = "" Or textfile = "All" Then Exit Sub
      Open App.Path & "\ListData\" & textfile & ".txt" For Input As #iFile
      Do While Not EOF(iFile)
         Line Input #iFile, textfile
         strArray = Split(textfile, ",")
         Y = ltbItem.ListItems.Count + 1
       For i = 0 To UBound(strArray) Step 2
         If Not Trim$(strArray(i)) = "" Then
            ltbItem.ListItems.Add (Y), , strArray(i)
            ltbItem.ListItems(Y).SubItems(1) = strArray(i + 1)
         End If
      Next i
         Loop
         Close #iFile
End Sub

Private Sub LoadcboUnits()   ' load the descriptors
   Dim textfile As String
   Dim strArrayUnit() As String
   Dim i As Integer
   Dim iFile As Integer
   
   cboUnits.Clear
   iFile = FreeFile
   Open App.Path & "\ListData\UnitDes.udr" For Input As #iFile
   Do While Not EOF(iFile)
      Line Input #iFile, textfile
      strArrayUnit = Split(textfile, ",")
      For i = 0 To UBound(strArrayUnit)
         If Not Trim$(strArrayUnit(i)) = "" Then
            cboUnits.AddItem strArrayUnit(i)
         End If
      Next i
      Loop
      Close #iFile
End Sub

Private Sub LoadcboCat()
   Dim strPath As String
   Dim fStg As String
   Dim fSLen As Integer
   'Load categories into combo box
   cboCat.AddItem "All"
   strPath = Dir$(App.Path & "\ListData\" & "*.txt")
   If Not strPath = "" Then                                    'yes, there are files here so
   Do                                                                  'go get them
      fSLen = Len(strPath) - 4                                'filename length minus extension
      fStg = Mid$(strPath, 1, fSLen)                        'filename without extension
      cboCat.AddItem fStg                                     'put filename into combobox
      strPath = Dir$
   Loop Until strPath = ""
Else
   MsgBox "No files found!", vbCritical + vbOKOnly, "File - Error"
End If
End Sub

Private Sub ImageUp_Click()
     txtQuan.Text = Val(txtQuan.Text) + 1
     If Val(txtQuan.Text) > 25 Then txtQuan.Text = "25"
     txtQuan.SelStart = 1                                     'keep carot of the right of number
End Sub

Private Sub ImageDown_Click()
    txtQuan.Text = Val(txtQuan.Text) - 1
    If Val(txtQuan.Text) < 1 Then txtQuan.Text = "1"
    txtQuan.SelStart = 1                                      'keep carot of the right of number
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Seld As String
   
   If Button = 2 Then                                           'show popup menu
   Seld = DPM1.Popup("Delete Item,Clear List,Edit List,Print List,-,Exit")
   Select Case Seld
      Case "Delete Item":
          cmdDeleteItem_Click
      Case "Clear List":
          cmdClear_Click
      Case "Edit List":
          cmdEdit_Click
          If cmdEdit.Caption = "Edit (Hide)" Then
             cmdEdit.Caption = "Edit (Show)"
         Else
             cmdEdit.Caption = "Edit (Hide)"
         End If
      Case "Print List":
          cmdPrintList_Click
      Case "Exit":
         cmdExit_Click
   End Select
End If
End Sub

Private Sub ltbItem_Click()
    If txtQuan.Text = "" Then txtQuan.Text = "1"
    lbltotUnitPrice.Caption = Val(txtQuan.Text) * ltbItem.SelectedItem.SubItems(1)
    lbltotUnitPrice.Caption = Format(lbltotUnitPrice, "##0.00")
End Sub

Private Sub ltbItem_DblClick()
   cmdADD_Click
End Sub

Private Sub txtEditCat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkTT.Value = 1 Then ShowToolTip txtEditCat.hwnd, "Click on a category to show current items.", "Categorys", Tip_Balloon, Tip_Info
End Sub

Private Sub txtEditDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkTT.Value = 1 Then ShowToolTip txtEditDesc.hwnd, "Add or Delete descriptors and Save." & vbCrLf & "Don't forget the comma.", "Edit Descriptors", Tip_Balloon, Tip_Info
End Sub

Private Sub txtEditItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If chkTT.Value = 1 Then ShowToolTip txtEditItems.hwnd, "Add or Delete items and price and press Save." & vbCrLf & "Format: Item,Price, (ex.item,0.00,)" & vbCrLf & "Don't forget the commas", "Edit Items and Price", Tip_Balloon, Tip_Info
End Sub

Private Sub txtNewCatName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If chkTT.Value = 1 Then ShowToolTip txtNewCatName.hwnd, "Add a New Category name here and press Save." & vbCrLf & "To delete, Select a category above and press Delete.", "Edit Category", Tip_Balloon, Tip_Info
End Sub

Private Sub txtQuan_GotFocus()
   txtQuan.SelStart = 2                                                'set carot to right of charactor
   If ListView1.ListItems.Count <> 0 Then ListView1.SelectedItem.Selected = False     'hide the selection bar (highlight)
End Sub

Private Sub txtQuan_KeyPress(KeyAscii As Integer)
    ' accept numbers and backspace only
    If InStr(1, "8 48 49 50 51 52 53 54 55 56 57", CStr(KeyAscii)) = 0 Then
         KeyAscii = 0
    End If
End Sub

Private Sub txtRecipeName_KeyPress(KeyAscii As Integer)
Dim lg As Integer
   If KeyAscii = 13 Then                                              'enter key was pressed
      rtfmain.Text = txtRecipeName.Text & vbCrLf & vbCrLf
      rtfmain.SetFocus
      lg = Len(rtfmain.Text)
      rtfmain.SelStart = lg
   End If
End Sub

Private Sub txtEditCat_Click()
'highlight category that was clicked on
SendKeys "{HOME}+{END}"
DoEvents
lblEditCatName.Caption = txtEditCat.SelText
EditLoadItems                                                           'load categories into textbox
End Sub

Private Sub ListViewSetup()
    Dim iFontHeight As Long
    Dim iBarHeight As Integer
  
    'Set up a few listview properties, these can be set on the property page instead of being listed here.
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Add 1, , "Qty"
    ListView1.ColumnHeaders(1).Width = 30
    ListView1.ColumnHeaders.Add 2, , "Descriptor"
    ListView1.ColumnHeaders(2).Width = 80
    ListView1.ColumnHeaders.Add 3, , "Item"
    ListView1.ColumnHeaders(3).Width = 165
    ListView1.ColumnHeaders.Add 4, , "Unit Price"
    ListView1.ColumnHeaders(4).Width = 60
    ListView1.ColumnHeaders.Add 5, , "Total Price"
    ListView1.ColumnHeaders(5).Width = 65
    ListView1.Width = ListView1.ColumnHeaders(1).Width + ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(3).Width + ListView1.ColumnHeaders(4).Width + ListView1.ColumnHeaders(5).Width
    ltbItem.ColumnHeaders.Add 1, , "Item"
    ltbItem.ColumnHeaders(1).Width = 145
    ltbItem.ColumnHeaders.Add 2, , "Price"
    ltbItem.ColumnHeaders(2).Width = 50
    ltbItem.Width = ltbItem.ColumnHeaders(1).Width + ltbItem.ColumnHeaders(2).Width + 20
    Me.ScaleMode = vbTwips 'make sure our form is In twips
    'Paints the green and white bars
    picGreenBar.ScaleMode = vbTwips
    picGreenBar.BorderStyle = vbBSNone 'this is important - we don't want To measure the border In our calcs.
    picGreenBar.AutoRedraw = True
    picGreenBar.Visible = False
    picGreenBar.Font = ListView1.Font
    picGreenBar.FontSize = ListView1.Font.Size
    iFontHeight = picGreenBar.TextHeight("b") + Screen.TwipsPerPixelY
    iBarHeight = (iFontHeight)
    picGreenBar.Width = ListView1.Width
    picGreenBar.Height = iBarHeight * 2
    picGreenBar.ScaleMode = vbUser
    picGreenBar.ScaleHeight = 2                              'bar-widths high
    picGreenBar.ScaleWidth = 1                               'bar-width wide
    picGreenBar.Line (0, 0)-(1, 1), vbWhite, BF           'white bars - modify vbWhite To change bar color
    picGreenBar.Line (0, 1)-(1, 2), &HF6F0E0, BF       'light green bars - modify RGB(x,x,x) To change bar color
   
    ListView1.PictureAlignment = lvwTile
    ListView1.Picture = picGreenBar.Image
End Sub

Private Sub EditLoadCat()
   Dim strPath As String
   Dim fStg As String
   Dim fSLen As Integer
   
   txtEditCat.Text = ""
   'Load categories into text box
   strPath = Dir(App.Path & "\ListData\" & "*.txt")
   If Not strPath = "" Then                                                 'yes, there are files here so
      Do                                                                           'go get them
         fSLen = Len(strPath) - 4                                         'filename length minus extension
         fStg = Mid$(strPath, 1, fSLen)                                 'filename without extension
         txtEditCat.Text = txtEditCat.Text & fStg & vbCrLf      'put filename into combobox
    
         strPath = Dir$
      Loop Until strPath = ""
   Else
      MsgBox "No files found!", vbCritical + vbOKOnly, "File - Error"
   End If
End Sub

Private Sub EditLoadItems()
   If lblEditCatName.Caption = "" Then
      MsgBox "Select a category to load", vbOKOnly, "No Category selected"
      txtEditItems.Text = ""
      Exit Sub
   End If
    LoadText (App.Path & "\ListData\" & lblEditCatName.Caption & ".txt"), txtEditItems
End Sub

Private Sub lstRecipe_Click()
   rtfmain.Text = ""                                                           'clears window before loading next recipe
   Call RTB_Load
End Sub

Private Sub List_Remove(TheList As ListBox)
   On Error Resume Next
   If TheList.ListCount < 0 Then Exit Sub
   TheList.RemoveItem TheList.ListIndex
End Sub

Public Function SaveText(strText As String, FileName As String) As Boolean
Dim iFile As Integer

On Error GoTo handle
    iFile = FreeFile
    Open FileName For Output As #iFile          'Opening the file to SaveText
        Print #iFile, strText                               'Printing  the text to the file
    Close #iFile                                              'Closing
    If FileExists(FileName) = False Then          'Check whether the file created
        MsgBox "Unexpectd error occured. File could not be saved", vbCritical, "Sorry"
        SaveText = False                                  'Returns 'False'
    Else
        SaveText = True                                    'Returns 'True'
    End If
Exit Function
handle:
    SaveText = False
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function

Private Sub LoadText(textfile As String, txtname As TextBox)
   Dim iFile As Integer
   
   txtname.Text = ""
    iFile = FreeFile
    Open textfile For Input As #iFile
        textfile = Input(LOF(iFile), iFile)
        txtname.Text = textfile
    Close #iFile
End Sub

Private Sub LoadTextInstr()                                            'load editing instructions
   txtInst.Text = "Descriptor Format:"
   txtInst.Text = txtInst.Text & " Max chars = 10 plus a comma on the end." & vbCrLf & vbCrLf
   txtInst.Text = txtInst.Text & "Add items and delete items in textboxes,"
   txtInst.Text = txtInst.Text & " then Press -Save- button." & vbCrLf & vbCrLf
   txtInst.Text = txtInst.Text & "For a new item, select a category, then type name comma price comma (ex. item,0.00,)."
   txtInst.Text = txtInst.Text & "Don't forget the commas, now Press -Save- button."
End Sub

Private Sub RTB_Save()
   Dim fFile As Integer
   
   fFile = FreeFile
   Open App.Path & "\RecipeFolder\" & txtRecipeName & ".rtf" For Output As fFile
   Print #fFile, txtRecipeName.Text & vbCrLf & vbCrLf & rtfmain.Text                                          ' String location you want To save
   Close fFile
End Sub

Private Sub RTB_Load()
   
   Dim FileLength As Integer
   Dim var1 As String
   Dim fFile As Integer
   
   fFile = FreeFile
   rtfmain.Text = ""
   'on startup load measurements into rtf box
   If firsttime = False Then
      Open App.Path & "\RecipeFolder\" & lstRecipe.Text & ".rtf" For Input As #fFile
   Else
      Open App.Path & "\RecipeFolder\" & "...Measurements.rtf" For Input As #fFile
   End If
      FileLength = LOF(fFile)
      var1 = Input(FileLength, #fFile)
      rtfmain.Text = var1
      rtfmain.SelStart = 0                                                'Puts Beginning of code at top
      Close #fFile
      firsttime = False
End Sub

Private Sub File_Delete(TList As ListBox)
   If TList = "" Then Exit Sub
   Kill App.Path & "\RecipeFolder\" & TList & ".rtf"
End Sub
