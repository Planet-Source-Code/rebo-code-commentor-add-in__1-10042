VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Settings"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtComPrefix 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1080
         TabIndex        =   3
         Text            =   "ŒŒŒ~~~"
         ToolTipText     =   "Single Comment Line prefix. Depends on fontsetting set in Editor."
         Top             =   1095
         Width           =   1020
      End
      Begin VB.TextBox txtName 
         Height          =   276
         Left            =   1080
         TabIndex        =   1
         Text            =   "User Name"
         ToolTipText     =   "Comment Block User Name."
         Top             =   525
         Width           =   2640
      End
      Begin VB.TextBox txtInitials 
         Height          =   288
         Left            =   1080
         TabIndex        =   2
         Text            =   "Initials"
         ToolTipText     =   "Single Comment Line User Initials."
         Top             =   810
         Width           =   1020
      End
      Begin VB.TextBox txtOrganisation 
         Height          =   276
         Left            =   1080
         TabIndex        =   0
         Text            =   "Organisation Name"
         ToolTipText     =   "Comment Block Organisation name"
         Top             =   240
         Width           =   2640
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Com. Prefix:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1155
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Initials:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Organisation:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   945
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   396
      Left            =   2940
      TabIndex        =   4
      ToolTipText     =   "Close Form"
      Top             =   1620
      Width           =   972
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtOrganisation.Text = ComOpt.Organisation
    txtInitials.Text = ComOpt.Initials
    txtName.Text = ComOpt.UserName
    txtComPrefix.Text = ComOpt.ComPrefix
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ComOpt.Organisation = txtOrganisation.Text
    ComOpt.UserName = txtName.Text
    ComOpt.Initials = txtInitials.Text
    ComOpt.ComPrefix = txtComPrefix.Text
    
    SaveSetting APP_CATEGORY, App.Title, "Organisation", txtOrganisation.Text
    SaveSetting APP_CATEGORY, App.Title, "UserName", txtName.Text
    SaveSetting APP_CATEGORY, App.Title, "Initials", txtInitials.Text
    SaveSetting APP_CATEGORY, App.Title, "ComPrefix", txtComPrefix.Text
End Sub



