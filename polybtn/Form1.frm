VERSION 5.00
Object = "*\AnPolBtn.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin nPolBtn.nPolyButton nPolyButton2 
      Height          =   1095
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "pentagon"
      Sides           =   5
   End
   Begin nPolBtn.nPolyButton nPolyButton1 
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "triangle"
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   600
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

