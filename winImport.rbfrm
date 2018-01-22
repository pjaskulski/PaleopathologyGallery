#tag Window
Begin Window winImport
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   False
   Composite       =   False
   Frame           =   1
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   5.81e+2
   ImplicitInstance=   False
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   False
   MaxWidth        =   32000
   MenuBar         =   ""
   MenuBarVisible  =   True
   MinHeight       =   470
   MinimizeButton  =   False
   MinWidth        =   1050
   Placement       =   0
   Resizeable      =   True
   Title           =   ""
   Visible         =   True
   Width           =   1.138e+3
   Begin Listbox lboxFiles
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   ""
      Border          =   True
      ColumnCount     =   2
      ColumnsResizable=   True
      ColumnWidths    =   "10%\r\n90%"
      DataField       =   ""
      DataSource      =   ""
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   ""
      EnableDragReorder=   ""
      GridLinesHorizontal=   0
      GridLinesVertical=   0
      HasHeading      =   True
      HeadingIndex    =   -1
      Height          =   443
      HelpTag         =   ""
      Hierarchical    =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   ""
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      RequiresSelection=   True
      Scope           =   0
      ScrollbarHorizontal=   ""
      ScrollBarVertical=   True
      SelectionType   =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   84
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   610
      _ScrollWidth    =   -1
   End
   Begin StaticText staticPath
      AutoDeactivate  =   True
      Bold            =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   20
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   True
      Multiline       =   ""
      Scope           =   0
      TabIndex        =   5
      TabPanelIndex   =   0
      Text            =   "Path"
      TextAlign       =   0
      TextColor       =   0
      TextFont        =   "System"
      TextSize        =   0
      Top             =   14
      Underline       =   ""
      Visible         =   True
      Width           =   46
   End
   Begin TextField txtPath
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
      InitialParent   =   ""
      Italic          =   ""
      Left            =   78
      LimitText       =   0
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   ""
      ReadOnly        =   True
      Scope           =   0
      TabIndex        =   6
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   0
      TextFont        =   "System"
      TextSize        =   0
      Top             =   12
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   806
   End
   Begin PushButton btnPath
      AutoDeactivate  =   True
      Bold            =   ""
      Cancel          =   ""
      Caption         =   "..."
      Default         =   ""
      Enabled         =   True
      Height          =   22
      HelpTag         =   "Change path..."
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   1078
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   12
      Underline       =   ""
      Visible         =   True
      Width           =   40
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
      Left            =   1038
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   True
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   4
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   539
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin PushButton btnImport
      AutoDeactivate  =   True
      Bold            =   ""
      Cancel          =   ""
      Caption         =   "Import"
      Default         =   ""
      Enabled         =   True
      Height          =   22
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   946
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   ""
      LockRight       =   True
      LockTop         =   ""
      Scope           =   0
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   0
      Top             =   539
      Underline       =   ""
      Visible         =   True
      Width           =   80
   End
   Begin ImageWell imgPreview
      AutoDeactivate  =   True
      Enabled         =   True
      Height          =   443
      HelpTag         =   ""
      Image           =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   644
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   8
      TabPanelIndex   =   0
      TabStop         =   True
      Top             =   84
      Visible         =   True
      Width           =   474
   End
   Begin TextField txtCount
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
      InitialParent   =   ""
      Italic          =   ""
      Left            =   967
      LimitText       =   0
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   ""
      ReadOnly        =   True
      Scope           =   0
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   False
      Text            =   ""
      TextColor       =   0
      TextFont        =   "System"
      TextSize        =   0
      Top             =   12
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   81
   End
   Begin StaticText staticImages
      AutoDeactivate  =   True
      Bold            =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   ""
      Left            =   910
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Multiline       =   ""
      Scope           =   0
      TabIndex        =   10
      TabPanelIndex   =   0
      Text            =   "Images:"
      TextAlign       =   0
      TextColor       =   0
      TextFont        =   "System"
      TextSize        =   0
      Top             =   14
      Underline       =   ""
      Visible         =   True
      Width           =   51
   End
   Begin ProgressBar ProgressImport
      AutoDeactivate  =   True
      Enabled         =   True
      Height          =   20
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   ""
      LockTop         =   ""
      Maximum         =   100
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   539
      Value           =   0
      Visible         =   False
      Width           =   610
   End
   Begin Thread ThreadImport
      Height          =   32
      Index           =   -2147483648
      Left            =   642
      LockedInPosition=   False
      Priority        =   5
      Scope           =   0
      StackSize       =   0
      TabPanelIndex   =   0
      Top             =   529
      Width           =   32
   End
   Begin TextField editDefaultSite
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
      InitialParent   =   ""
      Italic          =   ""
      Left            =   78
      LimitText       =   0
      LockBottom      =   ""
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   ""
      ReadOnly        =   ""
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   0
      TextFont        =   "System"
      TextSize        =   0
      Top             =   46
      Underline       =   ""
      UseFocusRing    =   True
      Visible         =   True
      Width           =   806
   End
   Begin StaticText staticDefaultSite
      AutoDeactivate  =   True
      Bold            =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   20
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
      Multiline       =   ""
      Scope           =   0
      TabIndex        =   12
      TabPanelIndex   =   0
      Text            =   "Site"
      TextAlign       =   0
      TextColor       =   0
      TextFont        =   "System"
      TextSize        =   0
      Top             =   46
      Underline       =   ""
      Visible         =   True
      Width           =   46
   End
   Begin Timer Timer1
      Height          =   32
      Index           =   -2147483648
      Left            =   682
      LockedInPosition=   False
      Mode            =   0
      Period          =   500
      Scope           =   0
      TabPanelIndex   =   0
      Top             =   529
      Width           =   32
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  lboxFiles.ColumnCount = 2
		  lboxFiles.ColumnType(0) = Listbox.TypeCheckbox
		  lboxFiles.ColumnType(1) = Listbox.TypeDefault
		  lboxFiles.ColumnWidths = "10%,90%"
		  
		  lboxFiles.Column(0).UserResizable = True
		  lboxFiles.Column(1).UserResizable = True
		  
		  lboxFiles.Heading(0) = "Import"
		  lboxFiles.Heading(1) = "File name"
		  
		  effect = New ScaleEffect
		  jpg = New JpegImporter
		  jpg_export = New JpegExporter
		  
		  Dim tmp As FolderItem
		  
		  #If DebugBuild Then
		    #If TargetWin32 Then
		      tmp = GetFolderItem("c:\PaleopathologyGallery\tmp\",FolderItem.PathTypeAbsolute)
		      Self.txtPath.Text = tmp.AbsolutePath
		    #Endif
		  #Endif
		  
		  txtCount.Text = "0"
		  Self.FileListRefresh = False
		  Self.StopImport = False
		  
		End Sub
	#tag EndEvent

	#tag Event
		Sub Resized()
		  ResizeControls()
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub InvertSelection()
		  // Invert the selection
		  
		  Dim i As Integer
		  
		  For i = 0 To lboxFiles.ListCount - 1
		    If lboxFiles.CellState(i,0) = Checkbox.CheckedStates.Checked Then
		      lboxFiles.CellCheck(i, 0) = False
		    Else
		      lboxFiles.CellCheck(i, 0) = True
		    End If
		  Next
		  
		  NumberOfSelected = lboxFiles.ListCount - NumberOfSelected
		  txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxFiles.ListCount)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ListDeselectAll()
		  Dim i As Integer
		  
		  For i = 0 To lboxFiles.ListCount - 1
		    lboxFiles.CellCheck(i, 0) = False
		  Next
		  
		  NumberOfSelected = 0
		  txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxFiles.ListCount)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ListSelectAll()
		  Dim i As Integer
		  
		  For i = 0 To lboxFiles.ListCount - 1
		    lboxFiles.CellCheck(i, 0) = True
		  Next
		  
		  NumberOfSelected = lboxFiles.ListCount
		  txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxFiles.ListCount)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ResizeControls()
		  Dim maxWidth, maxHeight As Integer
		  Dim factor As Double
		  
		  //txtPath.Width = staticImages.Left - txtPath.Left - 12
		  
		  If p <> NIL Then
		    maxWidth = Round(imgPreview.Width * App.ImageFactor)
		    maxHeight = Round(imgPreview.Height * App.ImageFactor)
		    
		    factor = Min( maxWidth / p.Width, maxHeight / p.Height )
		    factor = Min( factor, 1.0 ) // (don't scale it up if it's too small!)
		    imgPreview.Image = effect.BilinearScale(p,Round(p.Width*factor),Round(p.Height*factor))
		  End If
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		effect As ScaleEffect
	#tag EndProperty

	#tag Property, Flags = &h0
		fileList() As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		FileListRefresh As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		fnew As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		ImportValue As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		intCount As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		jpg As JpegImporter
	#tag EndProperty

	#tag Property, Flags = &h0
		jpg_export As JpegExporter
	#tag EndProperty

	#tag Property, Flags = &h0
		NumberOfSelected As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		p As Picture
	#tag EndProperty

	#tag Property, Flags = &h0
		StopImport As Boolean
	#tag EndProperty


#tag EndWindowCode

#tag Events lboxFiles
	#tag Event
		Sub Open()
		  #If TargetWin32
		    lboxFiles.TextSize=12
		    lboxFiles.TextFont="Tahoma"
		    lboxFiles.DefaultRowHeight = 20
		  #ElseIf TargetMacOS Then
		    lboxFiles.DefaultRowHeight = 28
		  #Endif
		End Sub
	#tag EndEvent
	#tag Event
		Function CellBackgroundPaint(g As Graphics, row As Integer, column As Integer) As Boolean
		  If row mod 2=0 then
		    g.foreColor=RGB(245,245,225)
		  else
		    g.foreColor=RGB(255,255,255)
		  end if
		  g.FillRect 0,0,g.width,g.height
		End Function
	#tag EndEvent
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  If Asc(key) = 32 Then
		    If lboxFiles.CellState(lboxFiles.ListIndex,0) = Checkbox.CheckedStates.Checked Then
		      lboxFiles.CellCheck(lboxFiles.ListIndex,0) = False
		      NumberOfSelected = NumberOfSelected - 1
		    Else
		      lboxFiles.CellCheck(lboxFiles.ListIndex,0) = True
		      NumberOfSelected = NumberOfSelected + 1
		    End If
		    txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxFiles.ListCount)
		  End If
		  
		  Return False
		End Function
	#tag EndEvent
	#tag Event
		Sub Change()
		  If Self.FileListRefresh = True Then Return
		  
		  If lboxFiles.ListCount = 0 Then Return
		  
		  Dim maxWidth, maxHeight As Integer
		  Dim imgWidth, imgHeight As Integer
		  Dim factor As Double
		  Dim f As FolderItem
		  
		  maxWidth = Round(imgPreview.Width * App.ImageFactor)
		  maxHeight = Round(imgPreview.Height * App.ImageFactor)
		  
		  f = lboxFiles.RowTag(lboxFiles.ListIndex)
		  
		  If f.Exists Then
		    p = jpg.OpenFromFile(f)
		    factor = Min( maxWidth / p.Width, maxHeight / p.Height )
		    factor = Min( factor, 1.0 )
		    If factor <> 1.0 Then
		      imgWidth = Round(p.Width*factor)
		      imgHeight = Round(p.Height*factor)
		      imgPreview.Image = effect.BilinearScale(p,imgWidth,imgHeight)
		    Else
		      imgPreview.Image = p
		    End If
		  Else
		    imgPreview.Image = NIL
		  End If
		  
		End Sub
	#tag EndEvent
	#tag Event
		Function ConstructContextualMenu(base as MenuItem, x as Integer, y as Integer) As Boolean
		  If Me.ListCount = 0 Then Return False
		  
		  base.Append(New MenuItem("Select all"))
		  base.Append(New MenuItem("Deselect all"))
		  base.Append(New MenuItem("Invert selection"))
		  
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Function ContextualMenuAction(hitItem as MenuItem) As Boolean
		  Select Case hititem.text
		  Case "Select all"
		    ListSelectAll()
		  Case "Deselect all"
		    ListDeselectAll()
		  Case "Invert selection"
		    InvertSelection
		  End Select
		  
		  Return True
		End Function
	#tag EndEvent
	#tag Event
		Sub CellAction(row As Integer, column As Integer)
		  If Column = 0 Then
		    If Me.CellState(row,column) = Checkbox.CheckedStates.Checked Then
		      NumberOfSelected = NumberOfSelected + 1
		    Else
		      NumberOfSelected = NumberOfSelected - 1
		    End If
		    txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxFiles.ListCount)
		  End If
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnPath
	#tag Event
		Sub Action()
		  If ThreadImport.State = Thread.Running Then
		    Return
		  End If
		  
		  Dim f, initFolder As FolderItem
		  Dim NumOfFiles, i As Integer
		  Dim dlg as New SelectFolderDialog
		  Dim fName As String
		  
		  dlg.ActionButtonCaption = "Select"
		  dlg.Title = "Title Property"
		  dlg.PromptText = "Prompt Text"
		  
		  If txtPath.Text <> "" Then
		    initFolder = GetFolderItem(txtPath.Text,FolderItem.PathTypeAbsolute)
		    dlg.InitialDirectory = initFolder
		  End If
		  
		  txtPath.SetFocus
		  f = dlg.ShowModal()
		  
		  Self.FileListRefresh = True
		  
		  If f <> NIL And f.Exists Then
		    txtPath.Text = f.AbsolutePath
		    
		    lboxFiles.DeleteAllRows
		    NumOfFiles = f.Count
		    
		    //lboxFiles.Enabled = False
		    
		    App.MouseCursor = System.Cursors.Wait
		    App.DoEvents
		    
		    For i = 1 to NumOfFiles
		      If f.Item(i).Directory = False Then
		        fname = f.Item(i).Name
		        If Right(fName,4) = ".jpg" Or Right(fName,5) = ".jpeg" Then
		          Self.lboxFiles.AddRow("")
		          Self.lboxFiles.ListIndex = Self.lboxFiles.LastIndex
		          Self.lboxFiles.RowTag(Self.lboxFiles.ListIndex) = f.Item(i)
		          Self.lboxFiles.CellCheck(Self.lboxFiles.ListIndex,0) = True
		          Self.lboxFiles.Cell(Self.lboxFiles.ListIndex,1) = fName
		        End if
		      End if
		    Next
		    
		    p = NIL
		    
		    //lboxFiles.Enabled = True
		    
		    App.MouseCursor = System.Cursors.StandardPointer
		    App.DoEvents
		    
		  End If
		  
		  NumberOfSelected = lboxFiles.ListCount
		  txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxFiles.ListCount)
		  
		  Self.FileListRefresh = False
		  lboxFiles.ListIndex = 0
		  lboxFiles.SetFocus
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnCancel
	#tag Event
		Sub Action()
		  
		  If ThreadImport.State = Thread.Running Then
		    Self.StopImport = True
		    Return
		  End If
		  
		  Self.imgPreview.Image = NIL
		  Self.Close
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events btnImport
	#tag Event
		Sub Action()
		  If ThreadImport.State = Thread.Running Then
		    Return
		  End If
		  
		  Dim i As Integer
		  
		  intCount = 0
		  
		  For i = 0 To Self.lboxFiles.ListCount - 1
		    If lboxFiles.CellState(i,0) = Checkbox.CheckedStates.Checked Then
		      intCount = intCount + 1
		    End If
		  Next
		  
		  If intCount = 0 Then
		    MsgBox("No file selected.")
		    Return
		  End If
		  
		  lboxFiles.Enabled = False
		  editDefaultSite.Enabled = False
		  
		  ProgressImport.Visible = True
		  ImportValue = 0
		  
		  App.DoEvents
		  
		  //Timer1.Mode = Timer.ModeMultiple
		  ThreadImport.Run
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ThreadImport
	#tag Event
		Sub Run()
		  Dim i, intCounter As Integer
		  Dim f, fs, CurItem As FolderItem
		  Dim rec As DatabaseRecord
		  Dim AddRowID As Integer
		  Dim maxWidth, maxHeight As Integer
		  Dim imgWidth, imgHeight As Integer
		  Dim factor As Double
		  Dim p_tmp, p As Picture
		  Dim rs As RecordSet
		  Dim fname As String
		  
		  f = SetPath("", "IMAGE", "image")
		  fs = SetPath("", "IMAGE", "s_image")
		  
		  intCounter = 0
		  
		  If intCount > 0 Then
		    For i = 0 To lboxFiles.ListCount - 1
		      If lboxFiles.CellState(i,0) = Checkbox.CheckedStates.Checked Then
		        CurItem = lboxFiles.RowTag(i)
		        rec = New DatabaseRecord
		        rec.Column("CAT_NO") = ""
		        rec.Column("SITE") = editDefaultSite.Text
		        rec.Column("GRAVE") = ""
		        rec.Column("BONE") = ""
		        rec.Column("SIDE") = ""
		        rec.Column("LOCATION") = ""
		        rec.Column("CHANGE_TYPE") = ""
		        rec.Column("PATH_TYPE") = ""
		        rec.Column("DIAGNOSIS") = ""
		        rec.Column("NOTES") = ""
		        
		        // Insert Record Into table "PALEO" of the Database
		        dbPaleo.InsertRecord "PALEO", rec
		        
		        // Display Error if one occurred
		        If dbPaleo.Error Then
		          DisplayDatabaseError True
		          Return
		        End if
		        
		        AddRowID = dbPaleo.lastRowID()
		        
		        fName = Str(AddRowID) + ".jpg"
		        If f.Child(fName).Exists Then
		          f.Child(fName).Delete
		        End If
		        
		        CurItem.CopyFileTo(f.Child(fName))
		        
		        maxWidth = 800
		        maxHeight = 600
		        
		        p = jpg.OpenFromFile(CurItem)
		        factor = Min(maxWidth / p.Width, maxHeight / p.Height)
		        factor = Min(factor, 1.0)
		        If factor <> 1.0 Then
		          imgWidth = Round(p.Width * factor)
		          imgHeight = Round(p.Height * factor)
		          p_tmp = effect.BilinearScale(p,imgWidth,imgHeight)
		        Else
		          imgWidth = Round(p.Width)
		          imgHeight = Round(p.Height)
		          p_tmp = p
		        End If
		        
		        jpg_export.SaveToFile(p_tmp,fs.Child("s_" + fName))
		        
		        // update SLIDE_NO field
		        rs = dbPaleo.SQLSelect("SELECT * FROM PALEO WHERE ID = " + Str(AddRowID))
		        If dbPaleo.Error Then
		          DisplayDatabaseError True
		          Return
		        End if
		        
		        rs.Edit //call Edit prior to modifiying the record and check for errors
		        If dbPaleo.Error then
		          DisplayDatabaseError True
		        End if
		        
		        // Update the record information
		        rs.Field("SLIDE_NO").StringValue = fName
		        rs.Field("ID_ZBIAUW").StringValue = "IA/" + App.Login + "/" + Format(rs.Field("ID").IntegerValue, "000000")
		        
		        // Update the record in the database
		        rs.Update
		        
		        If dbPaleo.Error then // handle errors
		          DisplayDatabaseError True
		          Return
		        End if
		        
		        NewImages.Value(rs.Field("ID_ZBIAUW").StringValue) =  AddRowID
		        
		        intCounter = intCounter + 1
		        
		      End If
		      
		      ImportValue = Round((intCounter / intCount) * 100)
		      ProgressImport.Value = ImportValue
		      If Self.StopImport = True Then
		        Exit
		      End If
		    Next
		    
		    // Commit changes to the database
		    dbPaleo.Commit
		  End If
		  
		  If Self.StopImport = False Then
		    Self.Close
		  Else
		    ProgressImport.Visible = False
		    btnImport.Enabled = True
		    btnCancel.Enabled = True
		    btnPath.Enabled = True
		    lboxFiles.Enabled = True
		    editDefaultSite.Enabled = True
		    Self.StopImport = False
		  End If
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events Timer1
	#tag Event
		Sub Action()
		  'If ThreadImport.State = Thread.Running Then
		  'ProgressImport.Value = ImportValue
		  'ProgressImport.Refresh
		  'Else
		  'Me.Mode = Timer.ModeOff
		  'winImport.Close
		  'End If
		End Sub
	#tag EndEvent
#tag EndEvents
