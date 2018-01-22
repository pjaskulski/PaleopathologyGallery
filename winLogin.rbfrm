#tag Window
Begin Window winLogin
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   1
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   133
   ImplicitInstance=   False
   LiveResize      =   False
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   ""
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   False
   MinWidth        =   64
   Placement       =   0
   Resizeable      =   False
   Title           =   "Login"
   Visible         =   True
   Width           =   328
   Begin GroupBox GroupBox1
      AutoDeactivate  =   True
      Bold            =   ""
      Caption         =   "Who are you?"
      Enabled         =   True
      Height          =   65
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   20
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   ""
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TextFont        =   "System"
      TextSize        =   0
      Top             =   14
      Underline       =   ""
      Visible         =   True
      Width           =   288
      Begin StaticText StaticText1
         AutoDeactivate  =   True
         Bold            =   ""
         DataField       =   ""
         DataSource      =   ""
         Enabled         =   True
         Height          =   20
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "GroupBox1"
         Italic          =   ""
         Left            =   29
         LockBottom      =   ""
         LockedInPosition=   False
         LockLeft        =   ""
         LockRight       =   ""
         LockTop         =   ""
         Multiline       =   ""
         Scope           =   0
         TabIndex        =   1
         TabPanelIndex   =   0
         Text            =   "The researcher's code"
         TextAlign       =   0
         TextColor       =   0
         TextFont        =   "System"
         TextSize        =   0
         Top             =   38
         Underline       =   ""
         Visible         =   True
         Width           =   130
      End
      Begin TextField editLogin
         AcceptTabs      =   ""
         Alignment       =   0
         AutoDeactivate  =   True
         BackColor       =   16777215
         Bold            =   ""
         Border          =   True
         CueText         =   ""
         DataField       =   ""
         DataSource      =   ""
         Enabled         =   True
         Format          =   ""
         Height          =   22
         HelpTag         =   ""
         Index           =   -2147483648
         InitialParent   =   "GroupBox1"
         Italic          =   ""
         Left            =   161
         LimitText       =   0
         LockBottom      =   ""
         LockedInPosition=   False
         LockLeft        =   ""
         LockRight       =   ""
         LockTop         =   ""
         Mask            =   ">"
         Password        =   ""
         ReadOnly        =   ""
         Scope           =   0
         TabIndex        =   2
         TabPanelIndex   =   0
         TabStop         =   True
         Text            =   ""
         TextColor       =   0
         TextFont        =   "System"
         TextSize        =   0
         Top             =   37
         Underline       =   ""
         UseFocusRing    =   True
         Visible         =   True
         Width           =   121
      End
   End
   Begin PushButton btnCancel
      AutoDeactivate  =   True
      Bold            =   ""
      Cancel          =   True
      Caption         =   "Cancel"
      Default         =   ""
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   228
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   ""
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   91
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin PushButton btnOK
      AutoDeactivate  =   True
      Bold            =   ""
      Cancel          =   ""
      Caption         =   "OK"
      Default         =   True
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   136
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   ""
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   91
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
End
#tag EndWindow

#tag WindowCode
#tag EndWindowCode

#tag Events editLogin
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  If Asc(key) = 13 Then
		    btnOK.SetFocus
		    Return True
		  End If
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events btnCancel
	#tag Event
		Sub Action()
		  App.Login = ""
		  Self.Close()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnOK
	#tag Event
		Sub Action()
		  
		  If Trim(editLogin.Text) <> "" Then
		    
		    App.Login = editLogin.Text
		    Self.Close()
		    
		  Else
		    
		    MsgBox("The researcher's code can not be empty.")
		    editLogin.SetFocus
		    Return
		    
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
