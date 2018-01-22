#tag Window
Begin Window winZoom
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   True
   Composite       =   False
   Frame           =   1
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   3.51e+2
   ImplicitInstance=   False
   LiveResize      =   False
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   True
   MaxWidth        =   32000
   MenuBar         =   ""
   MenuBarVisible  =   True
   MinHeight       =   64
   MinimizeButton  =   False
   MinWidth        =   64
   Placement       =   0
   Resizeable      =   False
   Title           =   "Untitled"
   Visible         =   True
   Width           =   4.7e+2
   Begin Canvas imgZoom
      AcceptFocus     =   True
      AcceptTabs      =   ""
      AutoDeactivate  =   True
      Backdrop        =   ""
      DoubleBuffer    =   True
      Enabled         =   True
      EraseBackground =   True
      Height          =   331
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   0
      UseFocusRing    =   True
      Visible         =   True
      Width           =   453
   End
   Begin ScrollBar ScrollBar2
      AcceptFocus     =   true
      AutoDeactivate  =   True
      Enabled         =   True
      Height          =   331
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   454
      LineStep        =   1
      LiveScroll      =   ""
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   True
      LockTop         =   True
      Maximum         =   100
      Minimum         =   0
      PageStep        =   20
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   0
      Value           =   0
      Visible         =   True
      Width           =   16
   End
   Begin ScrollBar ScrollBar1
      AcceptFocus     =   true
      AutoDeactivate  =   True
      Enabled         =   True
      Height          =   19
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   0
      LineStep        =   1
      LiveScroll      =   ""
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   ""
      Maximum         =   100
      Minimum         =   0
      PageStep        =   20
      Scope           =   0
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   332
      Value           =   0
      Visible         =   True
      Width           =   450
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Function CancelClose(appQuitting as Boolean) As Boolean
		  Self.jpg_pic = NIL
		  Self.buffer = NIL
		  
		  Return False
		End Function
	#tag EndEvent

	#tag Event
		Sub KeyUp(Key As String)
		  
		  If Asc(key) = 27 Then
		    Self.Close()
		  End If
		End Sub
	#tag EndEvent

	#tag Event
		Sub Open()
		  angle = 0
		  imgZoom.SetFocus
		End Sub
	#tag EndEvent

	#tag Event
		Sub Resized()
		  If buffer <> NIL Then
		    ScrollBar1.Maximum = buffer.Width - imgZoom.Width
		    ScrollBar2.Maximum = buffer.Height - imgZoom.Height
		    
		    If buffer.Width <= imgZoom.Width Then ScrollBar1.Enabled = False Else ScrollBar1.Enabled = True
		    If buffer.Height <= imgZoom.Height Then ScrollBar2.Enabled = False Else ScrollBar2.Enabled = True
		  End If
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub EditNotes()
		  Dim rs As RecordSet
		  Dim wn As New winNotes
		  
		  rs = dbPaleo.SQLSelect("SELECT NOTES, ID_ZBIAUW FROM PALEO WHERE ID = " + Str(App.CurRecID))
		  If dbPaleo.Error Then
		    DisplayDatabaseError True
		    Return
		  End if
		  
		  Self.ShowNotes = True
		  
		  rs.MoveFirst
		  wn.editNotes.Text = rs.Field("NOTES").StringValue
		  wn.strNotesBuff = rs.Field("NOTES").StringValue
		  wn.Title = "Notes: " + rs.Field("ID_ZBIAUW").StringValue
		  
		  #If TargetWin32 And UseEnableWin32 Then
		    ShowModally(wn, Self)
		  #Else
		    wn.ShowModalWithin(Self)
		  #EndIf
		  
		  Self.ShowNotes = False
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FullSize()
		  If Me.FullSize = True Then Return
		  
		  Dim rt As New RotateEffect
		  
		  buffer = jpg_pic
		  
		  factor = 1.0
		  
		  If angle <> 0 Then
		    buffer = rt.Apply(buffer, angle, RGB(255,255,255), RotationMethod.COLOR_WEIGHT)
		  End If
		  
		  ScrollBar1.Enabled = True
		  ScrollBar2.Enabled = True
		  
		  ScrollBar1.Maximum = buffer.Width - imgZoom.Width
		  ScrollBar2.Maximum = buffer.Height - imgZoom.Height
		  
		  Xvalue = 0
		  Yvalue = 0
		  Xscroll = 0
		  Yscroll = 0
		  MouseStartPosX = 0
		  MouseStartPosY = 0
		  
		  Me.FullSize = True
		  Me.FitSize = False
		  
		  Self.UpdateTitle
		  
		  imgZoom.Refresh
		  
		  ScrollBar1.Value = Xvalue
		  ScrollBar2.Value = Yvalue
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Rotate(strDirection As String)
		  Dim rt As New RotateEffect
		  Dim zoom_scale As New ScaleEffect
		  Dim maxWidth, maxHeight As Integer
		  
		  If strDirection = "LEFT" Then
		    angle = angle - 90
		  ElseIf strDirection = "RIGHT" Then
		    angle = angle + 90
		  Else
		    angle = 0
		  End If
		  
		  If angle >=360 Then
		    angle = 0
		  ElseIf angle = -90 Then
		    angle = 270
		  End If
		  
		  buffer = rt.Apply(jpg_pic, angle, RGB(255,255,255), RotationMethod.COLOR_WEIGHT)
		  
		  If Me.FullSize = False Then
		    maxWidth = Round(imgZoom.Width)
		    maxHeight = Round(imgZoom.Height)
		    
		    If Me.FitSize = True Then
		      factor = Min( maxWidth / buffer.Width, maxHeight / buffer.Height )
		      factor = Min( factor, 1.0 )
		    End If
		    
		    buffer = zoom_scale.BilinearScale(buffer,Round(buffer.Width*factor),Round(buffer.Height*factor))
		    
		    If Me.FitSize = False Then
		      ScrollBar1.Maximum = buffer.Width - imgZoom.Width
		      ScrollBar2.Maximum = buffer.Height - imgZoom.Height
		      
		      ScrollBar1.Value = 0
		      ScrollBar2.Value = 0
		    End If
		    
		    
		  Else
		    ScrollBar1.Maximum = buffer.Width - imgZoom.Width
		    ScrollBar2.Maximum = buffer.Height - imgZoom.Height
		    
		    ScrollBar1.Value = 0
		    ScrollBar2.Value = 0
		  End If
		  
		  Xvalue = 0
		  Yvalue = 0
		  MouseStartPosX = 0
		  MouseStartPosY = 0
		  
		  If Self.buffer.Width < Self.imgZoom.Width Then
		    Self.Xscroll = (Self.imgZoom.Width - Self.buffer.Width)/2
		  Else
		    Self.Xscroll = 0
		  End If
		  If Self.buffer.Height < Self.imgZoom.Height Then
		    Self.Yscroll = (Self.imgZoom.Height - Self.buffer.Height)/2
		  Else
		    Self.Yscroll = 0
		  End If
		  
		  imgZoom.Refresh
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScaleToValue(ZoomValue As Integer)
		  Dim maxWidth, maxHeight As Integer
		  Dim zoom_scale As New ScaleEffect
		  Dim rt As New RotateEffect
		  
		  buffer = jpg_pic
		  
		  maxWidth = Round(imgZoom.Width)
		  maxHeight = Round(imgZoom.Height)
		  
		  factor = ZoomValue/100 ' Min( maxWidth / jpg_pic.Width, maxHeight / jpg_pic.Height )
		  factor = Min( factor, 1.0 )
		  
		  buffer = zoom_scale.BilinearScale(buffer,Round(buffer.Width*factor),Round(buffer.Height*factor))
		  
		  If angle <> 0 Then
		    buffer = rt.Apply(buffer, angle, RGB(255,255,255), RotationMethod.COLOR_WEIGHT)
		  End If
		  
		  If buffer.Width > imgZoom.Width Then
		    ScrollBar1.Enabled = True
		  Else
		    ScrollBar1.Enabled = False
		  End if
		  
		  If buffer.Height > imgZoom.Height Then
		    ScrollBar2.Enabled = True
		  Else
		    ScrollBar2.Enabled = False
		  End If
		  
		  ScrollBar1.Maximum = buffer.Width - imgZoom.Width
		  ScrollBar2.Maximum = buffer.Height - imgZoom.Height
		  
		  Xvalue = 0
		  Yvalue = 0
		  MouseStartPosX = 0
		  MouseStartPosY = 0
		  
		  If Self.buffer.Width < Self.imgZoom.Width Then
		    Self.Xscroll = (Self.imgZoom.Width - Self.buffer.Width)/2
		  Else
		    Self.Xscroll = 0
		  End If
		  If Self.buffer.Height < Self.imgZoom.Height Then
		    Self.Yscroll = (Self.imgZoom.Height - Self.buffer.Height)/2
		  Else
		    Self.Yscroll = 0
		  End If
		  
		  Self.FullSize = False
		  Self.FitSize = False
		  
		  Self.UpdateTitle
		  
		  imgZoom.Refresh
		  
		  ScrollBar1.Value = Xvalue
		  ScrollBar2.Value = Yvalue
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ScaleToWin()
		  If Me.FitSize = True Then Return
		  
		  Dim rt As New RotateEffect
		  
		  Dim maxWidth, maxHeight As Integer
		  Dim zoom_scale As New ScaleEffect
		  
		  buffer = jpg_pic
		  
		  If angle <> 0 Then
		    buffer = rt.Apply(buffer, angle, RGB(255,255,255), RotationMethod.COLOR_WEIGHT)
		  End If
		  
		  maxWidth = Round(imgZoom.Width)
		  maxHeight = Round(imgZoom.Height)
		  
		  factor = Min( maxWidth / buffer.Width, maxHeight / buffer.Height )
		  factor = Min( factor, 1.0 )
		  
		  buffer = zoom_scale.BilinearScale(buffer,Round(buffer.Width*factor),Round(buffer.Height*factor))
		  
		  ScrollBar1.Enabled = False
		  ScrollBar2.Enabled = False
		  
		  ScrollBar1.Maximum = buffer.Width - imgZoom.Width
		  ScrollBar2.Maximum = buffer.Height - imgZoom.Height
		  
		  Xvalue = 0
		  Yvalue = 0
		  MouseStartPosX = 0
		  MouseStartPosY = 0
		  
		  If Self.buffer.Width < Self.imgZoom.Width Then
		    Self.Xscroll = (Self.imgZoom.Width - Self.buffer.Width)/2
		  Else
		    Self.Xscroll = 0
		  End If
		  If Self.buffer.Height < Self.imgZoom.Height Then
		    Self.Yscroll = (Self.imgZoom.Height - Self.buffer.Height)/2
		  Else
		    Self.Yscroll = 0
		  End If
		  
		  Self.FullSize = False
		  Self.FitSize = True
		  
		  Self.UpdateTitle
		  
		  imgZoom.Refresh
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UpdateTitle()
		  Self.Title = "ID: " + Self.ID + " - CAT_NO: " + Self.CAT_NO + " - SITE: " + Self.SITE + " - ZOOM: " + Str(Round(Self.factor * 100)) + "%"
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		angle As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		buffer As Picture
	#tag EndProperty

	#tag Property, Flags = &h0
		CAT_NO As String
	#tag EndProperty

	#tag Property, Flags = &h0
		factor As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		FitSize As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		FullSize As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		ID As String
	#tag EndProperty

	#tag Property, Flags = &h0
		jpg_pic As Picture
	#tag EndProperty

	#tag Property, Flags = &h0
		MouseStartPosX As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		MouseStartPosY As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ShowNotes As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		SITE As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Xscroll As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Xvalue As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Yscroll As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		Yvalue As Integer
	#tag EndProperty


#tag EndWindowCode

#tag Events imgZoom
	#tag Event
		Function ContextualMenuAction(hitItem as MenuItem) As Boolean
		  Select Case hititem.text
		  Case "Full size"
		    FullSize()
		  Case "90%"
		    ScaleToValue(90)
		  Case "80%"
		    ScaleToValue(80)
		  Case "70%"
		    ScaleToValue(70)
		  Case "60%"
		    ScaleToValue(60)
		  Case "50%"
		    ScaleToValue(50)
		  Case "40%"
		    ScaleToValue(40)
		  Case "30%"
		    ScaleToValue(30)
		  Case "20%"
		    ScaleToValue(20)
		  Case "10%"
		    ScaleToValue(10)
		  Case "Fit image to screen"
		    ScaleToWin()
		  Case "Save As..."
		    App.CopyFileTo()
		  Case "Close"
		    self.Close()
		  Case "Copy"
		    App.CopyToClipboard("IMAGE")
		  Case "Rotate right"
		    Rotate("RIGHT")
		  Case "Rotate left"
		    Rotate("LEFT")
		  Case "Notes"
		    EditNotes()
		  End select
		  
		  Return True
		  
		End Function
	#tag EndEvent
	#tag Event
		Function ConstructContextualMenu(base as MenuItem, x as Integer, y as Integer) As Boolean
		  If Self.ShowNotes Then Return False
		  
		  If base.Count > 0 Then
		    For i As Integer = base.Count To 1 Step -1
		      base.Remove(i - 1)
		    Next
		  End if
		  
		  base.Append(New MenuItem("Full size"))
		  base.Append(New MenuItem("90%"))
		  base.Append(New MenuItem("80%"))
		  base.Append(New MenuItem("70%"))
		  base.Append(New MenuItem("60%"))
		  base.Append(New MenuItem("50%"))
		  base.Append(New MenuItem("40%"))
		  base.Append(New MenuItem("30%"))
		  base.Append(New MenuItem("20%"))
		  base.Append(New MenuItem("10%"))
		  base.Append(New MenuItem("Fit image to screen"))
		  base.Append(New MenuBar("-"))
		  base.Append(New MenuItem("Save As..."))
		  base.Append(New MenuItem("Copy"))
		  base.Append(New MenuBar("-"))
		  base.Append(New MenuItem("Rotate right"))
		  base.Append(New MenuItem("Rotate left"))
		  base.Append(New MenuBar("-"))
		  base.Append(New MenuItem("Notes"))
		  base.Append(New MenuBar("-"))
		  base.Append(New MenuItem("Close"))
		  
		  Return True //display the contextual menu
		  
		End Function
	#tag EndEvent
	#tag Event
		Sub Paint(g As Graphics)
		  g.DrawPicture buffer,Xscroll,Yscroll
		End Sub
	#tag EndEvent
	#tag Event
		Sub KeyUp(Key As String)
		  If Asc(key) = 27 Then
		    Self.Close()
		  End If
		End Sub
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  If Self.FitSize = True Then Return True
		  
		  Select Case Asc(Key)
		  Case 31 'up arrow
		    ScrollBar2.Value = ScrollBar2.Value + ScrollBar2.LineStep
		  Case 29 'Right arrow
		    ScrollBar1.Value = ScrollBar1.Value + ScrollBar1.LineStep
		  Case 30 'Down arrow
		    ScrollBar2.Value = ScrollBar2.Value - ScrollBar2.LineStep
		  Case 28 'Left arrow
		    ScrollBar1.Value = ScrollBar1.Value - ScrollBar1.LineStep
		  Case 12 'Page Down
		    ScrollBar2.Value = ScrollBar2.Value + ScrollBar2.PageStep
		  Case 11 'Page Up
		    ScrollBar2.Value = ScrollBar2.Value - ScrollBar2.PageStep
		  End Select
		End Function
	#tag EndEvent
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  'Determine the starting coordinates
		  MouseStartPosX = X
		  MouseStartPosY = Y
		  
		  App.MouseCursor = System.Cursors.HandClosed
		  'Enable MouseDrag method
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Sub MouseDrag(X As Integer, Y As Integer)
		  Dim roznicaX, roznicaY As Integer
		  
		  'Update the canvas positions
		  roznicaX = X - MouseStartPosX
		  roznicaY = Y - MouseStartPosY
		  'The ending coordinate is now the starting position
		  MouseStartPosX = X
		  MouseStartPosY = Y
		  'Update the Canvas by change ScrollBars
		  
		  ScrollBar1.Value = ScrollBar1.Value - roznicaX
		  ScrollBar2.Value = ScrollBar2.Value - roznicaY
		  
		  
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub MouseUp(X As Integer, Y As Integer)
		  App.MouseCursor = System.Cursors.StandardPointer
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ScrollBar2
	#tag Event
		Sub ValueChanged()
		  Dim roznica As Integer
		  
		  roznica = Yvalue - Me.Value
		  Yvalue = Me.Value
		  
		  Yscroll = Yscroll + roznica
		  
		  imgZoom.Scroll 0,roznica
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ScrollBar1
	#tag Event
		Sub ValueChanged()
		  Dim roznica As Integer
		  
		  roznica = Xvalue - Me.Value
		  Xvalue = Me.Value
		  
		  Xscroll = Xscroll + roznica
		  
		  imgZoom.Scroll roznica,0
		End Sub
	#tag EndEvent
#tag EndEvents
