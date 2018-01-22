#tag Window
Begin Window winSync
   BackColor       =   16777215
   Backdrop        =   ""
   CloseButton     =   False
   Composite       =   False
   Frame           =   1
   FullScreen      =   False
   HasBackColor    =   False
   Height          =   581
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
   Width           =   1138
   Begin Listbox lboxRec
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
      Height          =   473
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
      Top             =   54
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
      Text            =   "Data folder"
      TextAlign       =   0
      TextColor       =   0
      TextFont        =   "System"
      TextSize        =   0
      Top             =   14
      Underline       =   ""
      Visible         =   True
      Width           =   63
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
      Left            =   95
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
      Width           =   789
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
      Height          =   473
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
      Top             =   54
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
      Text            =   "Records"
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
		  lboxRec.ColumnCount = 5
		  lboxRec.ColumnType(0) = Listbox.TypeCheckbox
		  lboxRec.ColumnType(1) = Listbox.TypeDefault
		  lboxRec.ColumnType(2) = Listbox.TypeDefault
		  lboxRec.ColumnType(3) = Listbox.TypeDefault
		  lboxRec.ColumnType(4) = Listbox.TypeDefault
		  
		  lboxRec.ColumnWidths = "10%,35%,35%,20%,0%"
		  
		  lboxRec.Column(0).UserResizable = True
		  lboxRec.Column(1).UserResizable = True
		  lboxRec.Column(2).UserResizable = True
		  lboxRec.Column(3).UserResizable = True
		  lboxRec.Column(4).UserResizable = False
		  
		  lboxRec.Heading(0) = "Import"
		  lboxRec.Heading(1) = "ID IAUW"
		  lboxRec.Heading(2) = "CAT. NO"
		  lboxRec.Heading(3) = "SITE"
		  lboxRec.Heading(4) = "IMAGE"
		  
		  effect = New ScaleEffect
		  jpg = New JpegImporter
		  jpg_export = New JpegExporter
		  
		  txtCount.Text = "0"
		  Self.FileListRefresh = False
		  Self.StopImport = False
		  
		  dbSync = New REALSQLDatabase
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
		  
		  For i = 0 To lboxRec.ListCount - 1
		    If lboxRec.CellState(i,0) = Checkbox.CheckedStates.Checked Then
		      lboxRec.CellCheck(i, 0) = False
		    Else
		      lboxRec.CellCheck(i, 0) = True
		    End If
		  Next
		  
		  NumberOfSelected = lboxRec.ListCount - NumberOfSelected
		  txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxRec.ListCount)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ListDeselectAll()
		  Dim i As Integer
		  
		  For i = 0 To lboxRec.ListCount - 1
		    lboxRec.CellCheck(i, 0) = False
		  Next
		  
		  NumberOfSelected = 0
		  txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxRec.ListCount)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ListSelectAll()
		  Dim i As Integer
		  
		  For i = 0 To lboxRec.ListCount - 1
		    lboxRec.CellCheck(i, 0) = True
		  Next
		  
		  NumberOfSelected = lboxRec.ListCount
		  txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxRec.ListCount)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ResizeControls()
		  Dim maxWidth, maxHeight As Integer
		  Dim factor As Double
		  
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
		dbFolder As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		dbSync As REALSQLDatabase
	#tag EndProperty

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

#tag Events lboxRec
	#tag Event
		Sub Open()
		  #If TargetWin32
		    Me.TextSize=12
		    Me.TextFont="Tahoma"
		    Me.DefaultRowHeight = 20
		  #ElseIf TargetMacOS Then
		    Me.DefaultRowHeight = 28
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
		    If Me.CellState(Me.ListIndex,0) = Checkbox.CheckedStates.Checked Then
		      Me.CellCheck(Me.ListIndex,0) = False
		      NumberOfSelected = NumberOfSelected - 1
		    Else
		      Me.CellCheck(Me.ListIndex,0) = True
		      NumberOfSelected = NumberOfSelected + 1
		    End If
		    txtCount.Text = Str(NumberOfSelected) + "/" + Str(Me.ListCount)
		  End If
		  
		  Return False
		End Function
	#tag EndEvent
	#tag Event
		Sub Change()
		  If Self.FileListRefresh = True Then Return
		  
		  If lboxRec.ListCount = 0 Then Return
		  
		  Dim maxWidth, maxHeight As Integer
		  Dim imgWidth, imgHeight As Integer
		  Dim factor As Double
		  Dim f As FolderItem
		  Dim strFileName As String
		  
		  maxWidth = Round(imgPreview.Width * App.ImageFactor)
		  maxHeight = Round(imgPreview.Height * App.ImageFactor)
		  
		  strFileName = lboxRec.Cell(lboxRec.ListIndex,4)
		  
		  If strFileName <> "" Then
		    strFileName = "s_" + strFileName
		    If Right(strFileName,4) <> ".jpg" Then
		      strFileName = strFileName + ".jpg"
		    End If
		    
		    f = dbFolder.Child("s_image").Child(strFileName)
		    
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
		    txtCount.Text = Str(NumberOfSelected) + "/" + Str(Me.ListCount)
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
		  Dim i As Integer
		  Dim dlg as New SelectFolderDialog
		  Dim rs As RecordSet
		  Dim strSQL As String
		  
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
		    
		    lboxRec.DeleteAllRows
		    
		    If Not f.Child("paleopathology.db").Exists Then
		      MsgBox "The selected folder does not contain a paleopathology database."
		      Return
		    End If
		    
		    If Not f.Child("image").Exists Then
		      MsgBox "The selected folder does not contain a image subfolder."
		      Return
		    End If
		    
		    If Not f.Child("s_image").Exists Then
		      MsgBox "The selected folder does not contain a s_image subfolder."
		      Return
		    End If
		    
		    App.MouseCursor = System.Cursors.Wait
		    App.DoEvents
		    
		    dbSync.DatabaseFile = f.Child("paleopathology.db")
		    dbFolder = f
		    
		    If dbSync.Connect = False Then
		      DisplayDatabaseError(False)
		      Return
		    End If
		    
		    strSQL = "select ID, ID_ZBIAUW, CAT_NO, SITE, SLIDE_NO from PALEO order by ID_ZBIAUW"
		    rs = dbSync.SQLSelect(strSQL)
		    lboxRec.DeleteAllRows
		    
		    While Not rs.eof
		      lboxRec.AddRow ""
		      lboxRec.RowTag(lboxRec.LastIndex) = rs.idxField(1).IntegerValue
		      lboxRec.CellCheck(lboxRec.LastIndex, 0) = True
		      For i = 2 To rs.fieldCount
		        lboxRec.cell(lboxRec.lastIndex, i - 1) = rs.idxField(i).stringValue
		      Next
		      
		      rs.moveNext
		    Wend
		    
		    rs.Close
		    
		    App.MouseCursor = System.Cursors.StandardPointer
		    App.DoEvents
		    
		  End If
		  
		  NumberOfSelected = lboxRec.ListCount
		  txtCount.Text = Str(NumberOfSelected) + "/" + Str(lboxRec.ListCount)
		  
		  Self.FileListRefresh = False
		  lboxRec.ListIndex = 0
		  lboxRec.SetFocus
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
		  
		  For i = 0 To lboxRec.ListCount - 1
		    If lboxRec.CellState(i,0) = Checkbox.CheckedStates.Checked Then
		      intCount = intCount + 1
		    End If
		  Next
		  
		  If intCount = 0 Then
		    MsgBox("No records selected.")
		    Return
		  End If
		  
		  lboxRec.Enabled = False
		  
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
		  Dim f, fs, SyncItem As FolderItem
		  Dim rec As DatabaseRecord
		  Dim rs As RecordSet
		  Dim AddRowID, SyncRecID As Integer
		  Dim maxWidth, maxHeight As Integer
		  Dim imgWidth, imgHeight As Integer
		  Dim factor As Double
		  Dim p_tmp, p As Picture
		  Dim rsSync As RecordSet
		  Dim fname As String
		  Dim strFileName As String
		  
		  f = SetPath("", "IMAGE", "image")
		  fs = SetPath("", "IMAGE", "s_image")
		  
		  intCounter = 0
		  
		  If intCount > 0 Then
		    For i = 0 To lboxRec.ListCount - 1
		      If lboxRec.CellState(i,0) = Checkbox.CheckedStates.Checked Then
		        SyncRecID = lboxRec.RowTag(i)
		        
		        rsSync = dbSync.SQLSelect("select ID_ZBIAUW,CAT_NO,SITE,GRAVE,BONE,SIDE,LOCATION,CHANGE_TYPE,PATH_TYPE,DIAGNOSIS,NOTES from PALEO where ID=" + Str(SyncRecID))
		        If Not rsSync.EOF Then
		          
		          rec = New DatabaseRecord
		          rec.Column("ID_ZBIAUW") = rsSync.idxField(1).stringValue
		          rec.Column("CAT_NO") = rsSync.idxField(2).stringValue
		          rec.Column("SITE") = rsSync.idxField(3).stringValue
		          rec.Column("GRAVE") = rsSync.idxField(4).stringValue
		          rec.Column("BONE") = rsSync.idxField(5).stringValue
		          rec.Column("SIDE") = rsSync.idxField(6).stringValue
		          rec.Column("LOCATION") = rsSync.idxField(7).stringValue
		          rec.Column("CHANGE_TYPE") = rsSync.idxField(8).stringValue
		          rec.Column("PATH_TYPE") = rsSync.idxField(9).stringValue
		          rec.Column("DIAGNOSIS") = rsSync.idxField(10).stringValue
		          rec.Column("NOTES") = rsSync.idxField(11).stringValue
		          
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
		          
		          strFileName = lboxRec.Cell(i,4)
		          If strFileName <> "" Then
		            
		            If Right(strFileName,4) <> ".jpg" Then
		              strFileName = strFileName + ".jpg"
		            End If
		            
		            // kopiowanie zdjęcia
		            SyncItem = dbFolder.Child("image").Child(strFileName)
		            If SyncItem.Exists Then
		              SyncItem.CopyFileTo(f.Child(fName))
		            End If
		            
		            // kopiowanie miniatury
		            SyncItem = dbFolder.Child("s_image").Child("s_" + strFileName)
		            If SyncItem.Exists Then
		              SyncItem.CopyFileTo(fs.Child("s_" + fName))
		            End If
		            
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
		            
		            // Update the record in the database
		            rs.Update
		            
		            If dbPaleo.Error then // handle errors
		              DisplayDatabaseError True
		              Return
		            End if
		            
		          End If
		          
		          // uzupełnienie słownika dodanych rekordów
		          NewImages.Value(lboxRec.Cell(i,1)) =  AddRowID
		          
		          // Commit changes to the database
		          dbPaleo.Commit
		          
		        End If
		        intCounter = intCounter + 1
		      End If
		      
		      ImportValue = Round((intCounter / intCount) * 100)
		      ProgressImport.Value = ImportValue
		      If Self.StopImport = True Then
		        Exit
		      End If
		    Next
		  End If
		  
		  If Self.StopImport = False Then
		    Self.Close
		  Else
		    ProgressImport.Visible = False
		    btnImport.Enabled = True
		    btnCancel.Enabled = True
		    btnPath.Enabled = True
		    lboxRec.Enabled = True
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
