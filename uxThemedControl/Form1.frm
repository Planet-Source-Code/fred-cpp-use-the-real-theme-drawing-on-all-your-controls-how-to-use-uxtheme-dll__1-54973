VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Themed Control Test"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin uxthemeTest.isThemedControl isThemedControl1 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
   End
   Begin VB.Label Label3 
      Caption         =   "By Fred.cpp"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Simple demonstration of a themed control, heavily commented"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "isThemed Control Template"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Fred.cpp 2004
' http://mx.geocities.com/fred_cpp
