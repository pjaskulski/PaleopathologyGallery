#tag Class
Protected Class App
Inherits Application
	#tag Event
		Sub Open()
		  jpg_reader = New JpegImporter
		  jpg_scale = New ScaleEffect
		  
		  dbPaleo = New REALSQLDatabase
		  NewImages = New Dictionary
		  
		  
		  // tworzenie katalogu na dane jeżeli uruchomiono samodzielny program, nie z IDE w trybie debugowania
		  #If Not DebugBuild Then
		    
		    Dim dataFolder, imageFolder, simageFolder As FolderItem
		    
		    dataFolder = GetFolderItem("").Child("data")
		    If Not dataFolder.Exists Then
		      dataFolder.CreateAsFolder
		    End If
		    
		    imageFolder = GetFolderItem("").Child("data").Child("image")
		    If Not imageFolder.Exists Then
		      imageFolder.CreateAsFolder
		    End If
		    
		    simageFolder = GetFolderItem("").Child("data").Child("s_image")
		    If Not simageFolder.Exists Then
		      simageFolder.CreateAsFolder
		    End If
		    
		  #EndIf
		  
		  dbPaleo.DatabaseFile = SetPath("paleopathology.db", "DB", "")
		  
		  Dim fBackup As FolderItem
		  fBackup = SetPath("paleopathology.backup", "BACKUP", "")
		  
		  If dbPaleo.DatabaseFile <> NIL And dbPaleo.DatabaseFile.Exists Then
		    Dim d As New Date
		    If fBackup.Exists Then
		      Dim md5DB, md5Back As String
		      
		      md5DB = GetMD5Hash(dbPaleo.DatabaseFile)
		      md5Back = GetMD5Hash(fBackup)
		      If md5DB <> md5Back Then
		        fBackup.Delete
		        dbPaleo.DatabaseFile.CopyFileTo(fBackup)
		        fBackup.ModificationDate = d
		      End If
		    Else
		      dbPaleo.DatabaseFile.CopyFileTo(fBackup)
		      fBackup.ModificationDate = d
		    End If
		  End If
		  
		  winMain.Maximize
		  
		  // read login
		  Dim file As FolderItem
		  file = SetPath("login.txt", "LOGIN", "")
		  
		  App.Login = ""
		  If file <> NIL And file.Exists Then
		    Dim readStream As TextInputStream
		    readStream = file.OpenAsTextFile
		    App.Login = readStream.ReadLine
		  End If
		  
		  // Set Researcher's code (login)
		  Dim wLogin As New winLogin
		  wLogin.editLogin.Text = App.Login
		  
		  #If TargetWin32 And UseEnableWin32 Then
		    ShowModally(wLogin, winMain)
		  #Else
		    wLogin.ShowModalWithin(winMain)
		  #EndIf
		  
		  If App.Login = "" Then
		    Quit
		  End If
		  
		  // save login
		  Dim fileStream As TextOutputStream
		  fileStream = file.CreateTextFile
		  fileStream.WriteLine App.Login
		  fileStream.Close
		  
		  If dbPaleo.DatabaseFile <> NIL And dbPaleo.DatabaseFile.Exists = True Then
		    If dbPaleo.Connect = False Then
		      DisplayDatabaseError(False)
		      Quit
		      Return
		    Else
		      UpdateDatabaseFile
		    End If
		  Else
		    CreateDatabaseFile
		  End If
		  
		  App.ImageFactor = 0.95
		  App.AutoQuit = True
		  App.UpdateDataList
		End Sub
	#tag EndEvent


	#tag MenuHandler
		Function CopyFileToItem() As Boolean Handles CopyFileToItem.Action
			App.CopyFileTo()
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function DbExport() As Boolean Handles DbExport.Action
			Dim curIndex As Integer
			Dim strRec, strSep As String
			Dim fileStream As TextOutputStream
			Dim file As FolderItem
			Dim txtType As New FileType
			
			strSep = ";"
			
			txtType.Name = "CSV/csv"
			txtType.MacType = "CSV"
			txtType.Extensions = "cvs"
			
			file = GetSaveFolderItem(txtType,"paleo.csv")
			If file <> NIL Then
			
			Try
			
			fileStream = file.CreateTextFile
			
			// 0 - CAT_NO varchar,
			// 1 - SITE varchar,
			// 2 - GRAVE varchar,
			// 3 - BONE varchar,
			// 4 - SIDE varchar,
			// 5 - LOCATION varchar,
			// 6 - CHANGE_TYPE varchar
			// 7 - PATH_TYPE varchar,
			// 8 - DIAGNOSIS varchar,
			// 9 - SLIDE_NO varchar,
			// 10 - NOTES varchar,
			// 11 - ID_ZBIAUW varchar
			
			strRec = "CAT_NO;ID;SITE;GRAVE/SKELETON;BONE;SIDE;LOCATION;CHANGE_TYPE;PATHOLOGY_TYPE;DIAGNOIS;FILE"
			fileStream.WriteLine strRec
			
			For curIndex = 0 To winMain.lboxDANE.ListCount - 1
			
			strRec = winMain.lboxDANE.Cell(curIndex,0) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,11) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,1) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,2) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,3) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,4) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,5) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,6) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,7) + strSep
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,8) + strSep
			If Right(winMain.lboxDANE.Cell(curIndex,9),4) <> ".jpg" Then
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,9) + ".jpg"
			Else
			strRec = strRec + winMain.lboxDANE.Cell(curIndex,9)
			End If
			fileStream.WriteLine strRec
			
			Next
			
			fileStream.Close
			
			Catch err As NilObjectException
			MsgBox "An unexpected problem occurred while exporting data."
			End Try
			
			End If
			
			Return True
			
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function DBSynchronization() As Boolean Handles DBSynchronization.Action
			
			//MsgBox "Funkcja w trakcie tworzenia..."
			App.SyncDatabase()
			Return True
			
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function DeleteItem() As Boolean Handles DeleteItem.Action
			App.DeleteRecord
			Return True
			
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function EditItem() As Boolean Handles EditItem.Action
			App.EditRecord
			Return True
			
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function FilterItem() As Boolean Handles FilterItem.Action
			App.SetFilter()
			
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function FindItem() As Boolean Handles FindItem.Action
			Dim wFind As New winFind
			
			#If TargetWin32 And UseEnableWin32 Then
			ShowModally(wFind, winMain)
			#Else
			wFind.ShowModalWithin(winMain)
			#EndIf
			
			winMain.Refresh
			
			App.FindRecord(False)
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function FindNextItem() As Boolean Handles FindNextItem.Action
			App.FindRecord(True)
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function HelpAbout() As Boolean Handles HelpAbout.Action
			MsgBox "PaleopathologyGallery v." + Str(App.MajorVersion) + "." + Str(App.MinorVersion) + "." + Str(App.NonReleaseVersion) _
			+ EndOfLine() + "Build: " + App.BuildDate.AbbreviatedDate
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function HelpHelp() As Boolean Handles HelpHelp.Action
			Dim f As FolderItem
			Dim strHTML As String
			
			f = SetPath("paleopathology.html","HELP","")
			strHTML = "file:///" + f.AbsolutePath
			
			ShowURL strHTML
			
			Return True
			
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function ImportItem() As Boolean Handles ImportItem.Action
			App.ImportFolder
			Return True
			
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function NewItem() As Boolean Handles NewItem.Action
			App.NewRecord
			Return True
			
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function ShowAllItem() As Boolean Handles ShowAllItem.Action
			App.UpdateDataList()
		End Function
	#tag EndMenuHandler

	#tag MenuHandler
		Function ZoomItem() As Boolean Handles ZoomItem.Action
			App.ShowZoom()
		End Function
	#tag EndMenuHandler


	#tag Method, Flags = &h0
		Sub CopyFileTo()
		  Dim strFile As String
		  Dim f, f_dest As FolderItem
		  
		  If Right(winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10),4) <> ".jpg" Then
		    strFile = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10) + ".jpg"
		  Else
		    strFile = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10)
		  End If
		  
		  f = SetPath(strFile, "IMAGE", "image")
		  
		  If f.Exists Then
		    f_dest = GetSaveFolderItem(FileTypesPaleo.Jpeg,f.Name)
		    If f_dest <> Nil Then
		      If f_dest.Name.Left(4) <> ".jpg" Then
		        f_dest.Name = f_dest.Name + ".jpg"
		      End If
		      f.CopyFileTo(f_dest)
		    End If
		  Else
		    MsgBox "File: " + strFile + " not exists"
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CopyToClipboard(strType As String)
		  Dim myClip As New Clipboard
		  Dim strFile As String
		  Dim fldPicture As FolderItem
		  Dim clipJPG As JpegImporter
		  Dim myPic As Picture
		  Dim strPre As String
		  
		  If strType = "IMAGE" Then
		    strPre = ""
		  ElseIf strType = "SIMAGE" Then
		    strPre = "s_"
		  End If
		  
		  clipJPG = New JpegImporter
		  If Right(winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10),4) <> ".jpg" Then
		    strFile = strPre + winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10) + ".jpg"
		  Else
		    strFile = strPre + winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10)
		  End If
		  
		  fldPicture = SetPath(strFile, "IMAGE", strPre + "image")
		  
		  If fldPicture.Exists Then
		    myClip.Text = ""
		    myPic = clipJPG.OpenFromFile(fldPicture)
		    myClip.Picture = myPic
		    If Not myClip.PictureAvailable Then
		      MsgBox "Unknown Clipboard error."
		    Else
		      Dim winMyInfo As New winInfo
		      winMyInfo.txtInfo.Text = "Copy to clipboard completed successfully."
		      winMyInfo.ShowModal
		    End If
		    myClip.Close
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CreateDatabaseFile()
		  // Create the database file
		  If dbPaleo.CreateDatabaseFile = False then
		    // Error While Creating the Database
		    MsgBox "Database Error" + EndOfLine + EndOfLine +_
		    "There was an error when creating the database."
		    Quit
		  End if
		  
		  // Create the tables and indexes for the database
		  dbPaleo.SQLExecute "create table PALEO (CAT_NO varchar,"+ _
		  "ID_ZBIAUW varchar, SITE varchar, GRAVE varchar, BONE varchar, "+ _
		  "SIDE varchar, LOCATION varchar, CHANGE_TYPE varchar, "+_
		  "PATH_TYPE varchar, DIAGNOSIS varchar, SLIDE_NO varchar, NOTES varchar, "+_
		  "ID integer NOT NULL PRIMARY KEY); create index ID_ZBIAUW on PALEO (ID_ZBIAUW); " +_
		  "create index CAT_NO on PALEO (CAT_NO); create index SITE on PALEO (SITE); "
		  
		  dbPaleo.Commit
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DeleteRecord()
		  // usuwanie rekordu
		  Dim d as New MessageDialog //declare the MessageDialog object
		  Dim b as MessageDialogButton //for handling the result
		  Dim rs as RecordSet
		  Dim recordID As Integer
		  Dim idx As Integer
		  Dim strRecordInfo As String
		  
		  d.Title = "Delete"
		  d.icon = MessageDialog.GraphicQuestion //display warning icon
		  d.ActionButton.Caption = "Delete"
		  d.CancelButton.Visible = True    //show the Cancel button
		  d.CancelButton.Default = True
		  strRecordInfo = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,0)
		  d.Message = "Do you want to delete current record (ID: "  + strRecordInfo + ")?"
		  d.Explanation = "Data and image will be lost. "
		  b = d.ShowModal //display the dialog
		  
		  Select Case b //determine which button was pressed
		  Case d.ActionButton
		    recordID = winMain.lboxDANE.RowTag(winMain.lboxDANE.ListIndex)
		    rs = dbPaleo.SQLSelect("SELECT * FROM PALEO WHERE ID = " + Str(recordID))
		    
		    If dbPaleo.Error Then
		      DisplayDatabaseError True
		      Return
		    End if
		    
		    rs.MoveFirst
		    If Not rs.EOF And Not rs.BOF Then
		      rs.DeleteRecord
		    End If
		    
		    If dbPaleo.Error then // handle errors
		      DisplayDatabaseError True
		      Return
		    End if
		    
		    dbPaleo.Commit
		    
		    idx = winMain.lboxDANE.ListIndex
		    winMain.lboxDANE.RemoveRow(idx)
		    If idx = 0 Then
		      winMain.lboxDANE.ListIndex = 0
		    Else
		      If idx > winMain.lboxDANE.ListCount - 1 Then
		        winMain.lboxDANE.ListIndex = idx - 1
		      Else
		        winMain.lboxDANE.ListIndex = idx
		      End If
		    End If
		  End select
		  
		  winMain.lboxDANE.SetFocus
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EditRecord()
		  Dim strFile As String
		  Dim maxWidth, maxHeight As Integer
		  Dim imgWidth, imgHeight As Integer
		  Dim factor As Double
		  Dim f As FolderItem
		  Dim jpg as JpegImporter
		  Dim effect As ScaleEffect
		  Dim p As Picture
		  Dim rs As RecordSet
		  Dim wEdit As New winEdit
		  
		  f = SetPath("", "IMAGE", "s_image")
		  
		  jpg = New JpegImporter
		  effect = New ScaleEffect
		  
		  wEdit.Hide
		  wEdit.Title = "Edit record"
		  
		  wEdit.editCAT_NO.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,1)
		  wEdit.cboSITE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,2)
		  wEdit.editGRAVE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,3)
		  wEdit.cboBONE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,4)
		  wEdit.cboSIDE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,5)
		  wEdit.editLOCATION.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,6)
		  wEdit.cboCHANGE_TYPE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,7)
		  wEdit.cboPATH_TYPE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,8)
		  wEdit.editDIAGNOSIS.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,9)
		  wEdit.editNOTES.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,11)
		  
		  // bufor rekordu
		  wEdit.buffCAT_NO = wEdit.editCAT_NO.Text
		  wEdit.buffSITE = wEdit.cboSITE.Text
		  wEdit.buffGRAVE = wEdit.editGRAVE.Text
		  wEdit.buffBONE = wEdit.cboBONE.Text
		  wEdit.buffSIDE = wEdit.cboSIDE.Text
		  wEdit.buffLOCATION = wEdit.editLOCATION.Text
		  wEdit.buffCHANGE_TYPE = wEdit.cboCHANGE_TYPE.Text
		  wEdit.buffPATH_TYPE = wEdit.cboPATH_TYPE.Text
		  wEdit.buffDIAGNOSIS = wEdit.editDIAGNOSIS.Text
		  wEdit.buffNOTES = wEdit.editNOTES.Text
		  
		  wEdit.RecordID = winMain.lboxDANE.RowTag(winMain.lboxDANE.ListIndex)
		  
		  If Right(winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10),4)<>".jpg" Then
		    strFile = "s_" + winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10) + ".jpg"
		  Else
		    strFile = "s_" + winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10)
		  End If
		  
		  maxWidth = Round(wEdit.imgFOTO.Width * App.ImageFactor)
		  maxHeight = Round(wEdit.imgFOTO.Height * App.ImageFactor)
		  
		  If f.Child(strFile).Exists Then
		    p = jpg.OpenFromFile(f.Child(strFile))
		    factor = Min( maxWidth / p.Width, maxHeight / p.Height )
		    factor = Min( factor, 1.0 )
		    imgWidth = Round(p.Width*factor)
		    imgHeight = Round(p.Height*factor)
		    wEdit.imgFOTO.Image = effect.BilinearScale(p,imgWidth,imgHeight)
		  Else
		    wEdit.imgFOTO.Image = NIL
		  End If
		  
		  wEdit.FotoName = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10)
		  
		  rs = dbPaleo.SQLSelect("SELECT ID_ZBIAUW FROM PALEO WHERE ID = " + Str(wEdit.RecordID))
		  If dbPaleo.Error Then
		    DisplayDatabaseError True
		    Return
		  Else
		    rs.MoveFirst
		    wEdit.Title = "ID: " + rs.Field("ID_ZBIAUW").StringValue
		  End if
		  
		  #If TargetWin32 And UseEnableWin32 Then
		    ShowModally(wEdit, winMain)
		  #Else
		    wEdit.ShowModalWithin(winMain)
		  #EndIf
		  
		  winMain.lboxDANE.SetFocus
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FindRecord(FromNext As Boolean)
		  Dim i,j, intStartFrom As Integer
		  Dim IsRecordFind As Boolean
		  
		  IsRecordFind = False
		  
		  If FromNext = True Then
		    intStartFrom = winMain.lboxDANE.ListIndex + 1
		  Else
		    intStartFrom = winMain.lboxDANE.ListIndex
		  End If
		  
		  If winMain.MyFindText <> "" Then
		    For i = intStartFrom To winMain.lboxDANE.ListCount - 1
		      For j = 0 To winMain.lboxDANE.ColumnCount - 1
		        If winMain.lboxDANE.Cell(i,j).InStr(winMain.MyFindText) <> 0  Then
		          winMain.lboxDANE.SetFocus
		          winMain.lboxDANE.ListIndex = i
		          winMain.lboxDANE.Selected(i)= True
		          IsRecordFind  = True
		          Exit
		        End If
		      Next
		      If IsRecordFind  = True Then Exit
		    Next
		    If IsRecordFind  = False Then
		      MsgBox "Cannot find the string: " + winMain.MyFindText
		    End If
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ImportFolder()
		  //import grupy fotografii z folderu
		  Dim wImport As New winImport
		  
		  wImport.Hide
		  NewImages.Clear
		  wImport.Title = "Import"
		  
		  #If TargetWin32 And UseEnableWin32 Then
		    ShowModally(wImport, winMain)
		  #Else
		    wImport.ShowModalWithin(winMain)
		  #EndIf
		  
		  //winImport.Timer1.Mode = Timer.ModeOff
		  
		  App.RefreshListAfterImport
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub NewRecord(DropImage As FolderItem = NIL)
		  // Metoda dodaje nowy rekord do bazy danych
		  Dim wEdit As New winEdit
		  
		  wEdit.Title = "New record"
		  wEdit.RecordID = 0
		  
		  wEdit.editCAT_NO.Text = ""
		  wEdit.cboSITE.Text = ""
		  wEdit.editGRAVE.Text = ""
		  wEdit.cboBONE.Text = ""
		  wEdit.cboSIDE.Text = ""
		  wEdit.editNOTES.Text = ""
		  wEdit.cboCHANGE_TYPE.Text = ""
		  wEdit.cboPATH_TYPE.Text = ""
		  wEdit.editDIAGNOSIS.Text = ""
		  wEdit.editNOTES.Text = ""
		  
		  // bufor rekordu
		  wEdit.buffCAT_NO = ""
		  wEdit.buffSITE = ""
		  wEdit.buffGRAVE = ""
		  wEdit.buffBONE = ""
		  wEdit.buffSIDE = ""
		  wEdit.buffLOCATION = ""
		  wEdit.buffCHANGE_TYPE = ""
		  wEdit.buffPATH_TYPE = ""
		  wEdit.buffDIAGNOSIS = ""
		  wEdit.buffNOTES = ""
		  
		  If DropImage <> NIL Then
		    wEdit.GetImage(DropImage)
		  End If
		  
		  #If TargetWin32 And UseEnableWin32 Then
		    ShowModally(wEdit, winMain)
		  #Else
		    wEdit.ShowModalWithin(winMain)
		  #EndIf
		  
		  winMain.lboxDANE.SetFocus
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub PopulateListBox(lb As ListBox, rs As RecordSet)
		  // Populate the specified listbox with the data in the recordset
		  // that is provided.  This will loop through the records in the
		  // recordset and add rows to the listbox that contain the data
		  // from the recordset.
		  
		  Dim i As Integer
		  
		  lb.deleteAllRows
		  
		  // Loop until we reach the end of the recordset
		  while not rs.eof
		    lb.addRow ""// add a new row
		    lb.RowTag(lb.lastIndex) = rs.idxField(1).IntegerValue
		    // Loop through all (after first - ID field) of the fields in the recordset and add the data to the correct
		    // column in the listbox
		    for i = 2 to rs.fieldCount
		      // The listbox Cell property is 0-based so we need to subtract 2 from the database field
		      // number to get the correct correct column number.  This means field 2 is in column 0 of the listbox.
		      lb.cell( lb.lastIndex, i-2 ) = rs.idxField( i ).stringValue
		    next
		    
		    rs.moveNext// move to the next record
		  wend
		  
		  
		  // If the listbox is set to be sorted by a particular column then we want to
		  // sort the listbox contents after we populate it, so that they appear in the
		  // correct order.
		  if lb.sortedColumn > -1 then// the listbox is sorted by a column
		    lb.sort// sort the listbox data using the current sort settings
		  end
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RefreshImageMain(indexDane as Integer)
		  Dim maxWidth, maxHeight As Integer
		  Dim imgWidth, imgHeight As Integer
		  Dim strFile As String
		  Dim fs As FolderItem
		  Dim p as Picture
		  Dim jpg As JpegImporter
		  Dim factor As Double
		  Dim effect As ScaleEffect
		  
		  jpg = new JpegImporter
		  effect = new ScaleEffect
		  
		  fs = SetPath("", "IMAGE", "s_image")
		  
		  maxWidth = Round(winMain.imgFOTO.Width * App.ImageFactor)
		  maxHeight = Round(winMain.imgFOTO.Height * App.ImageFactor)
		  
		  If Right(winMain.lboxDANE.Cell(indexDane,10),4) <> ".jpg" Then
		    strFile = "s_" + winMain.lboxDANE.Cell(indexDane,10) + ".jpg"
		  Else
		    strFile = "s_" + winMain.lboxDANE.Cell(indexDane,10)
		  End If
		  
		  If fs.Child(strFile).Exists Then
		    p = jpg.OpenFromFile(fs.Child(strFile))
		    factor = Min( maxWidth / p.Width, maxHeight / p.Height )
		    factor = Min( factor, 1.0 )
		    If factor <> 1.0 Then
		      imgWidth = Round(p.Width * factor)
		      imgHeight = Round(p.Height * factor)
		      winMain.imgFOTO.Image = effect.BilinearScale(p,imgWidth,imgHeight)
		    Else
		      winMain.imgFOTO.Image = p
		    End If
		  Else
		    winMain.imgFOTO.Image = NIL
		  End If
		  
		  winMain.FotoName = winMain.lboxDANE.Cell(indexDane,10)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RefreshListAfterImport()
		  Dim rs As RecordSet
		  
		  For Each key As Variant In NewImages.Keys
		    rs = dbPaleo.SQLSelect("SELECT * FROM PALEO WHERE ID = " + Str(NewImages.Value(key)))
		    If dbPaleo.Error Then
		      DisplayDatabaseError True
		      Return
		    End if
		    rs.MoveFirst
		    
		    winMain.lboxDANE.AddRow("")
		    winMain.lboxDANE.ListIndex = winMain.lboxDANE.LastIndex
		    winMain.lboxDANE.RowTag(winMain.lboxDANE.ListIndex) = NewImages.Value(key)
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,0) = key
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,1) = ""
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,2) = rs.Field("SITE").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10) = Str(NewImages.Value(key)) + ".jpg"
		  Next
		  
		  winMain.lboxDANE.SetFocus
		  
		  App.RefreshImageMain(winMain.lboxDANE.ListIndex)
		  
		  winMain.GroupBox1.Caption = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,0)
		  winMain.editCAT_NO.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,1)
		  winMain.editSITE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,2)
		  winMain.editGRAVE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,3)
		  winMain.editBONE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,4)
		  winMain.editSIDE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,5)
		  winMain.editLOCATION.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,6)
		  winMain.editCHANGE_TYPE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,7)
		  winMain.editPATH_TYPE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,8)
		  winMain.editDIAGNOSIS.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,9)
		  winMain.editNOTES.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,11)
		  
		  App.UpdateTitle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RefreshStaticField(indexDane As Integer)
		  winMain.GroupBox1.Caption = winMain.lboxDANE.Cell(indexDane,0)
		  winMain.editCAT_NO.Text = winMain.lboxDANE.Cell(indexDane,1)
		  winMain.editSITE.Text = winMain.lboxDANE.Cell(indexDane,2)
		  winMain.editGRAVE.Text = winMain.lboxDANE.Cell(indexDane,3)
		  winMain.editBONE.Text = winMain.lboxDANE.Cell(indexDane,4)
		  winMain.editSIDE.Text = winMain.lboxDANE.Cell(indexDane,5)
		  winMain.editLOCATION.Text = winMain.lboxDANE.Cell(indexDane,6)
		  winMain.editCHANGE_TYPE.Text = winMain.lboxDANE.Cell(indexDane,7)
		  winMain.editPATH_TYPE.Text = winMain.lboxDANE.Cell(indexDane,8)
		  winMain.editDIAGNOSIS.Text = winMain.lboxDANE.Cell(indexDane,9)
		  winMain.editNOTES.Text = winMain.lboxDANE.Cell(indexDane,11)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetFilter()
		  Dim winF As New winFilter
		  
		  winF.editFILTER_CAT_NO.Text = ""
		  winF.editFILTER_GRAVE.Text = ""
		  winF.editFILTER_LOCATION.Text = ""
		  winF.cboFILTER_SIDE.Text = ""
		  winF.editFILTER_DIAGNOSIS.Text = ""
		  winF.editFILTER_NOTES.Text = ""
		  winF.cboFILTER_BONE.Text = ""
		  winF.cboFILTER_CHANGE_TYPE.Text = ""
		  winF.cboFILTER_PATH_TYPE.Text = ""
		  winF.cboFILTER_SITE.Text = ""
		  
		  #If TargetWin32 And UseEnableWin32 Then
		    ShowModally(winF, winMain)
		  #Else
		    winF.ShowModalWithin(winMain)
		  #EndIf
		  
		  If App.MyQueryPar.RunQuery = True Then
		    App.UpdateDataListQuery()
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ShowZoom()
		  Dim strFile As String
		  Dim f As FolderItem
		  Dim strCatNo, strSite, strDiagnosis, strID As String
		  
		  If winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10)<>"" Then
		    If Right(winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10),4)<>".jpg" Then
		      strFile = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10) + ".jpg"
		    Else
		      strFile = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10)
		    End If
		  Else
		    strFile = ""
		  End If
		  
		  f = SetPath(strFile, "IMAGE", "image")
		  
		  If strFile <> "" Then
		    If f.Exists Then
		      Dim winZ As New winZoom
		      winZ.Hide
		      
		      App.MouseCursor = System.Cursors.Wait
		      App.DoEvents
		      
		      winZ.Maximize
		      winZ.jpg_pic = App.jpg_reader.OpenFromFile(f)
		      
		      Dim maxWidth, maxHeight As Integer
		      
		      maxWidth = Round(winZ.imgZoom.Width)
		      maxHeight = Round(winZ.imgZoom.Height)
		      
		      winZ.factor = Min( maxWidth / winZ.jpg_pic.Width, maxHeight / winZ.jpg_pic.Height )
		      winZ.factor = Min( winZ.factor, 1.0 )
		      winZ.buffer = App.jpg_scale.BilinearScale(winZ.jpg_pic,Round(winZ.jpg_pic.Width*winZ.factor),Round(winZ.jpg_pic.Height*winZ.factor))
		      
		      winZ.ScrollBar1.Maximum = winZ.buffer.Width - winZ.imgZoom.Width
		      winZ.ScrollBar2.Maximum = winZ.buffer.Height - winZ.imgZoom.Height
		      
		      If winZ.buffer.Width <= winZ.imgZoom.Width Then winZ.ScrollBar1.Enabled = False Else winZ.ScrollBar1.Enabled = True
		      If winZ.buffer.Height <= winZ.imgZoom.Height Then winZ.ScrollBar2.Enabled = False Else winZ.ScrollBar2.Enabled = True
		      
		      winZ.Xvalue = 0
		      winZ.Yvalue = 0
		      If winZ.buffer.Width < winZ.imgZoom.Width Then
		        winZ.Xscroll = (winZ.imgZoom.Width - winZ.buffer.Width)/2
		      Else
		        winZ.Xscroll = 0
		      End If
		      If winZ.buffer.Height < winZ.imgZoom.Height Then
		        winZ.Yscroll = (winZ.imgZoom.Height - winZ.buffer.Height)/2
		      Else
		        winZ.Yscroll = 0
		      End If
		      
		      winZ.FullSize = False
		      winZ.FitSize = True
		      winZ.angle = 0
		      
		      #If TargetWin32
		        winZ.imgZoom.EraseBackground = False
		        winZ.imgZoom.DoubleBuffer = True
		      #ElseIf TargetMacOS
		        winZ.imgZoom.DoubleBuffer = False
		        winZ.imgZoom.EraseBackground = True
		      #ElseIf TargetLinux
		        winZ.imgZoom.DoubleBuffer = False
		        winZ.imgZoom.EraseBackground = True
		      #Endif
		      
		      winZ.CAT_NO = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,1)
		      winZ.SITE = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,2)
		      winZ.ID = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,0)
		      winZ.UpdateTitle
		      
		      App.MouseCursor = System.Cursors.StandardPointer
		      App.DoEvents
		      
		      #If TargetWin32 And UseEnableWin32 Then
		        ShowModally(winZ, winMain)
		      #Else
		        winZ.ShowModalWithin(winMain)
		      #EndIf
		      
		    Else
		      MsgBox "File " + strFile + " not exists"
		    End If
		  Else
		    MsgBox("The current record does not contain an image.")
		  End If
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SyncDatabase()
		  //import rekordów z innej bazy danych
		  
		  Dim i As Integer
		  Dim rs As RecordSet
		  Dim wSync As New winSync
		  
		  NewImages.Clear
		  
		  wSync.Hide
		  wSync.Title = "Synchronization"
		  
		  #If TargetWin32 And UseEnableWin32 Then
		    ShowModally(wSync, winMain)
		  #Else
		    wSync.ShowModalWithin(winMain)
		  #EndIf
		  
		  //winImport.Timer1.Mode = Timer.ModeOff
		  
		  For Each key As Variant In NewImages.Keys
		    rs = dbPaleo.SQLSelect("SELECT * FROM PALEO WHERE ID = " + Str(NewImages.Value(key)))
		    If dbPaleo.Error Then
		      DisplayDatabaseError True
		      Return
		    End if
		    rs.MoveFirst
		    
		    winMain.lboxDANE.AddRow("")
		    winMain.lboxDANE.ListIndex = winMain.lboxDANE.LastIndex
		    winMain.lboxDANE.RowTag(winMain.lboxDANE.ListIndex) = NewImages.Value(key)
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,0) = key
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,1) = rs.Field("CAT_NO").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,2) = rs.Field("SITE").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,3) = rs.Field("GRAVE").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,4) = rs.Field("BONE").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,5) = rs.Field("SIDE").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,6) = rs.Field("LOCATION").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,7) = rs.Field("CHANGE_TYPE").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,8) = rs.Field("PATH_TYPE").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,9) = rs.Field("DIAGNOSIS").StringValue
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,10) = Str(NewImages.Value(key)) + ".jpg"
		    winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,11) = rs.Field("NOTES").StringValue
		  Next
		  
		  winMain.lboxDANE.SetFocus
		  
		  App.RefreshImageMain(winMain.lboxDANE.ListIndex)
		  
		  winMain.GroupBox1.Caption = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,0)
		  winMain.editCAT_NO.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,1)
		  winMain.editSITE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,2)
		  winMain.editGRAVE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,3)
		  winMain.editBONE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,4)
		  winMain.editSIDE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,5)
		  winMain.editLOCATION.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,6)
		  winMain.editCHANGE_TYPE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,7)
		  winMain.editPATH_TYPE.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,8)
		  winMain.editDIAGNOSIS.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,9)
		  winMain.editNOTES.Text = winMain.lboxDANE.Cell(winMain.lboxDANE.ListIndex,11)
		  
		  App.UpdateTitle
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UpdateDatabaseFile()
		  Dim rs, idx as RecordSet
		  Dim dbColumnExist As Boolean
		  
		  // Pole ID_ZBIAUW
		  dbColumnExist = False
		  
		  rs = dbPaleo.SQLSelect("PRAGMA table_info(PALEO)")
		  If rs.RecordCount > 0 Then
		    rs.MoveFirst
		    While Not rs.eof
		      If rs.Field("name").StringValue = "ID_ZBIAUW" Then
		        dbColumnExist = True
		      End If
		      rs.MoveNext
		    Wend
		    rs.Close
		  End If
		  
		  If Not dbColumnExist Then
		    dbPaleo.SQLExecute "alter table PALEO add column ID_ZBIAUW varchar"
		    dbPaleo.Commit
		  End If
		  
		  // uzupełnienie pola ID_ZBIAUW w instniejących rekordach
		  rs = dbPaleo.SQLSelect("select * from PALEO")
		  
		  If rs.RecordCount > 0 Then
		    rs.MoveFirst
		    While Not rs.EOF
		      If rs.Field("ID_ZBIAUW").StringValue = "" Then
		        
		        rs.Edit //call Edit prior to modifiying the record and check for errors
		        If dbPaleo.Error Then
		          DisplayDatabaseError True
		          Return
		        Else
		          rs.Field("ID_ZBIAUW").StringValue = "IA/" + App.Login + "/" + Format(rs.Field("ID").IntegerValue, "000000")
		          rs.Update
		          If dbPaleo.Error then // handle errors
		            DisplayDatabaseError True
		            Return
		          End if
		        End If
		        
		      ElseIf Left(rs.Field("ID_ZBIAUW").StringValue,4) = "ZBIA" Then
		        
		        rs.Edit //call Edit prior to modifiying the record and check for errors
		        If dbPaleo.Error Then
		          DisplayDatabaseError True
		          Return
		        Else
		          rs.Field("ID_ZBIAUW").StringValue = Replace(rs.Field("ID_ZBIAUW").StringValue,"ZBIA","IA")
		          rs.Update
		          If dbPaleo.Error then // handle errors
		            DisplayDatabaseError True
		            Return
		          End if
		        End If
		        
		      End If
		      
		      rs.MoveNext
		    Wend
		  End If
		  rs.Close
		  
		  dbPaleo.Commit
		  
		  // index ID_ZBIAUW
		  dbColumnExist = False
		  idx = dbPaleo.IndexSchema("PALEO")
		  If idx<> NIL And idx.RecordCount > 0 Then
		    idx.MoveFirst
		    While Not idx.eof
		      If idx.Field("IndexName").StringValue = "ID_ZBIAUW" Then
		        dbColumnExist = True
		      End If
		      idx.MoveNext
		    Wend
		    idx.Close
		  End If
		  
		  If Not dbColumnExist Then
		    dbPaleo.SQLExecute "create index ID_ZBIAUW on PALEO (ID_ZBIAUW) "
		    dbPaleo.Commit
		  End If
		  
		  // index CAT_NO
		  dbColumnExist = False
		  idx = dbPaleo.IndexSchema("PALEO")
		  If idx<> NIL And idx.RecordCount > 0 Then
		    idx.MoveFirst
		    While Not idx.eof
		      If idx.Field("IndexName").StringValue = "CAT_NO" Then
		        dbColumnExist = True
		      End If
		      idx.MoveNext
		    Wend
		    idx.Close
		  End If
		  
		  If Not dbColumnExist Then
		    dbPaleo.SQLExecute "create index CAT_NO on PALEO (CAT_NO) "
		    dbPaleo.Commit
		  End If
		  
		  // index SITE
		  dbColumnExist = False
		  idx = dbPaleo.IndexSchema("PALEO")
		  If idx<> NIL And idx.RecordCount > 0 Then
		    idx.MoveFirst
		    While Not idx.eof
		      If idx.Field("IndexName").StringValue = "SITE" Then
		        dbColumnExist = True
		      End If
		      idx.MoveNext
		    Wend
		    idx.Close
		  End If
		  
		  If Not dbColumnExist Then
		    dbPaleo.SQLExecute "create index SITE on PALEO (SITE) "
		    dbPaleo.Commit
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub UpdateDataList()
		  Dim sql as String
		  Dim rs as recordSet
		  
		  // Build the SQL statement that will be used to select the records
		  sql = "SELECT ID, ID_ZBIAUW, CAT_NO, SITE, GRAVE, BONE,"+_
		  "SIDE, LOCATION, CHANGE_TYPE, PATH_TYPE, DIAGNOSIS, SLIDE_NO, NOTES FROM PALEO"
		  
		  // Now we select the records from the database and add them to the list.
		  rs = dbPaleo.SQLSelect(sql)
		  PopulateListBox winMain.lboxDANE, rs
		  App.RecCount = rs.RecordCount
		  rs.Close()
		  winMain.lboxDANE.SetFocus
		  winMain.lboxDANE.ListIndex = 0
		  winMain.Title = "PaleopathologyGallery [All records ("+ Str(App.RecCount) +")]"
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UpdateDataListQuery()
		  Dim sql, sql_add as String
		  Dim rs as recordSet
		  Dim QueryRecCount As Integer
		  
		  sql = "SELECT ID, ID_ZBIAUW, CAT_NO, SITE, GRAVE, BONE,"+_
		  "SIDE, LOCATION, CHANGE_TYPE, PATH_TYPE, DIAGNOSIS, SLIDE_NO, NOTES FROM PALEO"
		  
		  If App.MyQueryPar.CAT_NO <>"" Or App.MyQueryPar.SITE <>"" Or App.MyQueryPar.GRAVE <> "" Or _
		    App.MyQueryPar.BONE <> "" Or App.MyQueryPar.SIDE <> "" Or App.MyQueryPar.LOCATION <> "" Or _
		    App.MyQueryPar.CHANGE_TYPE <> "" Or App.MyQueryPar.PATH_TYPE <> "" Or App.MyQueryPar.DIAGNOSIS <> "" Or _
		    App.MyQueryPar.SLIDE_NO <> "" Then
		    sql = sql + " WHERE"
		    
		    If App.MyQueryPar.CAT_NO <>"" Then
		      If App.MyQueryPar.ex_CAT_NO = True Then
		        sql_add = " CAT_NO = '" + App.MyQueryPar.CAT_NO + "'"
		      Else
		        sql_add = " CAT_NO LIKE '%" + App.MyQueryPar.CAT_NO + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.SITE <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_SITE = True Then
		        sql_add = " SITE = '" + App.MyQueryPar.SITE +"'"
		      Else
		        sql_add = " SITE LIKE '%" + App.MyQueryPar.SITE + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.GRAVE <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_GRAVE = True Then
		        sql_add = " GRAVE = '" + App.MyQueryPar.GRAVE + "'"
		      Else
		        sql_add = " GRAVE LIKE '%" + App.MyQueryPar.GRAVE + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.BONE <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_BONE = True Then
		        sql_add = " BONE = '" + App.MyQueryPar.BONE + "'"
		      Else
		        sql_add = " BONE LIKE '%" + App.MyQueryPar.BONE + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.SIDE <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_SIDE = True Then
		        sql_add = " SIDE = '" + App.MyQueryPar.SIDE + "'"
		      Else
		        sql_add = " SIDE LIKE '%" + App.MyQueryPar.SIDE + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.LOCATION <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_LOCATION = True Then
		        sql_add = " LOCATION = '" + App.MyQueryPar.LOCATION + "'"
		      Else
		        sql_add = " LOCATION LIKE '%" + App.MyQueryPar.LOCATION + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.CHANGE_TYPE <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_CHANGE_TYPE = True Then
		        sql_add = " CHANGE_TYPE = '" + App.MyQueryPar.CHANGE_TYPE + "'"
		      Else
		        sql_add = " CHANGE_TYPE LIKE '%" + App.MyQueryPar.CHANGE_TYPE + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.PATH_TYPE <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_PATH_TYPE = True Then
		        sql_add = " PATH_TYPE = '" + App.MyQueryPar.PATH_TYPE + "'"
		      Else
		        sql_add = " PATH_TYPE LIKE '%" + App.MyQueryPar.PATH_TYPE + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.DIAGNOSIS <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_DIAGNOSIS = True Then
		        sql_add = " DIAGNOSIS = '" + App.MyQueryPar.DIAGNOSIS + "'"
		      Else
		        sql_add = " DIAGNOSIS LIKE '%" + App.MyQueryPar.DIAGNOSIS + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		    
		    If App.MyQueryPar.SLIDE_NO <>"" Then
		      If Right(sql,5) <> "WHERE" Then
		        sql = sql + " AND"
		      End If
		      If App.MyQueryPar.ex_SLIDE_NO = True Then
		        sql_add = " SLIDE_NO = '" + App.MyQueryPar.SLIDE_NO + "'"
		      Else
		        sql_add = " SLIDE_NO LIKE '%" + App.MyQueryPar.SLIDE_NO + "%'"
		      End If
		      sql = sql + sql_add
		    End If
		  End If
		  
		  'MsgBox sql
		  
		  rs = dbPaleo.SQLSelect(sql)
		  PopulateListBox winMain.lboxDANE, rs
		  QueryRecCount = rs.RecordCount
		  rs.Close()
		  winMain.lboxDANE.SetFocus
		  winMain.lboxDANE.ListIndex = 0
		  winMain.Title = "PaleopathologyGallery [Filter: (" + Str(QueryRecCount) + " records from " + str(App.RecCount) + ")]"
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UpdateTitle()
		  Dim sql as String
		  Dim rs as recordSet
		  
		  If winMain.Title.InStr("Filter") > 0 Then
		    winMain.Title = winMain.Title + " [new records added]"
		  Else
		    // Build the SQL statement that will be used to select the records
		    sql = "SELECT ID FROM PALEO"
		    rs = dbPaleo.SQLSelect(sql)
		    App.RecCount = rs.RecordCount
		    rs.Close()
		    winMain.Title = "PaleopathologyGallery [All records ("+ Str(App.RecCount) +")]"
		  End If
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		CurRecID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ImageFactor As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		jpg_reader As JpegImporter
	#tag EndProperty

	#tag Property, Flags = &h0
		jpg_scale As ScaleEffect
	#tag EndProperty

	#tag Property, Flags = &h0
		Login As String
	#tag EndProperty

	#tag Property, Flags = &h0
		mCurrentModalHandle As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		MyQueryPar As QueryStructure
	#tag EndProperty

	#tag Property, Flags = &h0
		RecCount As Integer
	#tag EndProperty


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant

	#tag Constant, Name = kHelpAbout, Type = String, Dynamic = False, Default = \"About", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"About Ortner"
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="RecCount"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ImageFactor"
			Group="Behavior"
			InitialValue="0"
			Type="Double"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Login"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CurRecID"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="mCurrentModalHandle"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
