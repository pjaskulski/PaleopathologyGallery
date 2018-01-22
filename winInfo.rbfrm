#tag Window
Begin Window winInfo
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   False
   Composite       =   False
   Frame           =   11
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   60
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   ""
   MenuBarVisible  =   False
   MinHeight       =   64
   MinimizeButton  =   False
   MinWidth        =   64
   Placement       =   0
   Resizeable      =   False
   Title           =   "Info"
   Visible         =   True
   Width           =   300
   Begin Timer TimerInfo
      Enabled         =   True
      Height          =   32
      Index           =   -2147483648
      Left            =   0
      LockedInPosition=   False
      Mode            =   2
      Period          =   2500
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   0
      Visible         =   True
      Width           =   32
   End
   Begin StaticText txtInfo
      AutoDeactivate  =   True
      Bold            =   True
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   39
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
      Multiline       =   True
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextAlign       =   0
      TextColor       =   &h000000
      TextFont        =   "System"
      TextSize        =   0
      Top             =   12
      Underline       =   ""
      Visible         =   True
      Width           =   260
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  TimerInfo.Reset
		End Sub
	#tag EndEvent


#tag EndWindowCode

#tag Events TimerInfo
	#tag Event
		Sub Action()
		  Self.Close
		End Sub
	#tag EndEvent
#tag EndEvents
