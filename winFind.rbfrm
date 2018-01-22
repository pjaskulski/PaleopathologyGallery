#tag Window
Begin Window winFind
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   1
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   142
   ImplicitInstance=   False
   LiveResize      =   False
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   ""
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   True
   MinWidth        =   64
   Placement       =   0
   Resizeable      =   False
   Title           =   "Find record"
   Visible         =   True
   Width           =   405
   Begin GroupBox GroupBox1
      AutoDeactivate  =   True
      Bold            =   ""
      Caption         =   ""
      Enabled         =   True
      Height          =   68
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   20
      LockBottom      =   ""
      LockedInPosition=   True
      LockLeft        =   ""
      LockRight       =   ""
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   14
      Underline       =   ""
      Visible         =   True
      Width           =   365
      Begin TextField editFind
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
         Left            =   31
         LimitText       =   0
         LockBottom      =   ""
         LockedInPosition=   True
         LockLeft        =   ""
         LockRight       =   ""
         LockTop         =   ""
         Mask            =   ""
         Password        =   ""
         ReadOnly        =   ""
         Scope           =   0
         TabIndex        =   0
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
         Width           =   343
      End
   End
   Begin PushButton btnOK
      AutoDeactivate  =   True
      Bold            =   ""
      Cancel          =   ""
      Caption         =   "Find"
      Default         =   ""
      Enabled         =   True
      Height          =   28
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   213
      LockBottom      =   ""
      LockedInPosition=   True
      LockLeft        =   ""
      LockRight       =   ""
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   94
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin PushButton btnCancel
      AutoDeactivate  =   True
      Bold            =   ""
      Cancel          =   True
      Caption         =   "Cancel"
      Default         =   ""
      Enabled         =   True
      Height          =   28
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   305
      LockBottom      =   ""
      LockedInPosition=   True
      LockLeft        =   ""
      LockRight       =   ""
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   94
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  #If TargetMacOS Then
		    btnOK.Height = 20
		    btnCancel.Height = 20
		  #EndIf
		End Sub
	#tag EndEvent


#tag EndWindowCode

#tag Events editFind
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  If Asc(Key) = 13 Then
		    btnOk.SetFocus
		    Return True
		  End If
		  
		End Function
	#tag EndEvent
#tag EndEvents
#tag Events btnOK
	#tag Event
		Sub Action()
		  If Trim(editFind.Text) <> "" Then
		    winMain.MyFindText = Trim(editFind.Text)
		    Self.Close()
		  Else
		    MsgBox "Text to find is empty!"
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnCancel
	#tag Event
		Sub Action()
		  winMain.MyFindText = ""
		  Self.Close()
		End Sub
	#tag EndEvent
#tag EndEvents
