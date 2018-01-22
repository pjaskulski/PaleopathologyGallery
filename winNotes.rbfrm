#tag Window
Begin Window winNotes
   BackColor       =   &hFFFFFF
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   1
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   3.05e+2
   ImplicitInstance=   False
   LiveResize      =   False
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   ""
   MenuBarVisible  =   True
   MinHeight       =   100
   MinimizeButton  =   False
   MinWidth        =   200
   Placement       =   1
   Resizeable      =   True
   Title           =   "Notes"
   Visible         =   True
   Width           =   5.58e+2
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
      Left            =   458
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   True
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   265
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin PushButton btnOK
      AutoDeactivate  =   True
      Bold            =   ""
      Cancel          =   ""
      Caption         =   "Save"
      Default         =   True
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   366
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   True
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   265
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin TextArea editNotes
      AcceptTabs      =   ""
      Alignment       =   0
      AutoDeactivate  =   True
      BackColor       =   &hFFFFFF
      Bold            =   ""
      Border          =   True
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   239
      HelpTag         =   ""
      HideSelection   =   True
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   20
      LimitText       =   ""
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Multiline       =   True
      ReadOnly        =   ""
      Scope           =   0
      ScrollbarHorizontal=   ""
      ScrollbarVertical=   True
      Styled          =   False
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &h000000
      TextFont        =   "System"
      TextSize        =   0
      Top             =   14
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   518
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Function CancelClose(appQuitting as Boolean) As Boolean
		  Dim n as Integer
		  Dim blSave As Boolean
		  
		  If editNotes.Text <> strNotesBuff Then
		    n = MsgBox("The data in the Notes field has changed. Close without saving changes?",36)
		    
		    If n = 7 Then
		      blSave = True
		    Else
		      blSave = False
		    End If
		  End If
		  
		  Return blSave
		End Function
	#tag EndEvent


	#tag Property, Flags = &h0
		strNotesBuff As String
	#tag EndProperty


#tag EndWindowCode

#tag Events btnCancel
	#tag Event
		Sub Action()
		  
		  Self.Close()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnOK
	#tag Event
		Sub Action()
		  Dim rs As RecordSet
		  
		  rs = dbPaleo.SQLSelect("SELECT * FROM PALEO WHERE ID = " + Str(App.CurRecID))
		  If dbPaleo.Error Then
		    DisplayDatabaseError True
		    Return
		  End if
		  
		  rs.Edit //call Edit prior to modifiying the record and check for errors
		  If dbPaleo.Error then
		    DisplayDatabaseError True
		  End if
		  
		  // PALEO Table fields:
		  // SELECT ID, CAT_NO, SITE, GRAVE, BONE, SIDE, LOCATION, CHANGE_TYPE, PATH_TYPE, DIAGNOSIS, SLIDE_NO, NOTES
		  // Update the record information
		  
		  rs.Field("NOTES").StringValue = editNOTES.Text
		  
		  // Update the record in the database
		  rs.Update
		  If dbPaleo.Error then // handle errors
		    DisplayDatabaseError True
		    Return
		  End if
		  
		  // Commit changes to the database
		  dbPaleo.Commit
		  
		  // Update the current row of listbox
		  winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,11) = editNOTES.Text
		  winMain.editNOTES.Text = editNOTES.Text
		  
		  strNotesBuff = editNOTES.Text
		  
		  Self.Close()
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events editNotes
	#tag Event
		Function ConstructContextualMenu(base as MenuItem, x as Integer, y as Integer) As Boolean
		  If base.Count > 0 Then
		    For i As Integer = base.Count To 1 Step -1
		      base.Remove(i - 1)
		    Next
		  End if
		  
		  base.Append( New MenuItem("Cut") )
		  base.Append( New MenuItem("Copy") )
		  base.Append( New MenuItem("Paste") )
		  base.Append( New MenuItem("Delete") )
		  base.Append( New MenuItem( MenuItem.TextSeparator ))
		  base.Append( New MenuItem("Select All") )
		  base.Append( New MenuItem("Deselect All") )
		  
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Function ContextualMenuAction(hitItem as MenuItem) As Boolean
		  Select Case hititem.text
		  Case "Cut"
		    Me.Copy
		    Me.SelText = ""
		  Case "Copy" 
		    Me.Copy
		  Case "Paste"
		    Me.Paste
		  Case "Delete"
		    Me.SelText = ""
		  Case "Select All"
		    Me.SelectAll
		  Case "Deselect All"
		    Me.SelLength = 0
		  End select
		  
		  Return True
		End Function
	#tag EndEvent
#tag EndEvents
