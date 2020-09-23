VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test App"
   ClientHeight    =   4980
   ClientLeft      =   3000
   ClientTop       =   2505
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6015
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   7
      Left            =   3480
      Max             =   100
      TabIndex        =   16
      Top             =   3360
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   6
      Left            =   3480
      Max             =   100
      TabIndex        =   15
      Top             =   3000
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   5
      Left            =   3480
      Max             =   100
      TabIndex        =   14
      Top             =   2640
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   4
      Left            =   3480
      Max             =   100
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   3
      Left            =   3480
      Max             =   100
      TabIndex        =   12
      Top             =   1800
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   2
      Left            =   3480
      Max             =   100
      TabIndex        =   11
      Top             =   1200
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      Left            =   3480
      Max             =   100
      TabIndex        =   10
      Top             =   600
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      Left            =   3480
      Max             =   100
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin Project1.tjProgress tjProgress1 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
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
   Begin Project1.tjProgress tjProgress1 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColor    =   255
      CaptionYPosition=   0
      ProgressType    =   1
   End
   Begin Project1.tjProgress tjProgress1 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionXPosition=   1
      ProgressType    =   2
   End
   Begin Project1.tjProgress tjProgress1 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionXPosition=   2
      CaptionYPosition=   2
      ProgressType    =   3
   End
   Begin Project1.tjProgress tjProgress1 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCaption     =   0   'False
      ProgressType    =   4
   End
   Begin Project1.tjProgress tjProgress1 
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionType     =   1
      Caption         =   "My Progress Bar"
      ProgressType    =   5
   End
   Begin Project1.tjProgress tjProgress1 
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ProgressType    =   6
   End
   Begin Project1.tjProgress tjProgress1 
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ProgressType    =   7
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HScroll1_Change(Index As Integer)
    tjProgress1(Index).Value = HScroll1(Index).Value

End Sub
