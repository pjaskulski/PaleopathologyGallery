#tag Module
Protected Module mMain
	#tag Method, Flags = &h0
		Sub DisplayDatabaseError(doRollback as Boolean)
		  // Display a dialog that shows information about the error that
		  // has occurred from the database engine. If requested a Rollback
		  // will also happen on the database in order to undo any changes
		  // that happened since the last commit.
		  MsgBox "Database Error: " + str(dbPaleo.ErrorCode) + EndOfLine +_
		  EndOfLine + dbPaleo.ErrorMessage
		  // Rollback changes to the database if specified
		  If doRollback then
		    dbPaleo.Rollback
		  End if
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetMD5Hash(f As FolderItem) As String
		  // test parameters
		  If f = Nil Then Return ""
		  If Not f.Exists Then Return ""
		  If f.Directory Then Return ""
		  
		  // read file f to a string
		  // Note: RB string supports any binary data
		  Dim InputData As TextInputStream
		  Dim FileData As String
		  Dim MD5Dgt As String
		  Dim m as New MD5Digest
		  
		  InputData = f.OpenAsTextFile
		  
		  If InputData = Nil Then
		    // file f is not readable
		    Return ""
		  Else
		    FileData = InputData.ReadAll
		    InputData.Close
		  End If
		  
		  // get the MD5 digest
		  m.Process(FileData)
		  MD5Dgt = m.Value
		  
		  Return MD5Dgt
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SetPath(strFile As String, strType As String, strFolder As String) As FolderItem
		  Dim f As FolderItem
		  
		  If strType = "IMAGE" And strFile <> "" Then
		    
		    #If DebugBuild Then
		      #If TargetWin32 Then
		        f = GetFolderItem("c:\PaleopathologyGallery\data\" + strFolder + "\" + strFile, FolderItem.PathTypeAbsolute)
		      #ElseIf TargetLinux Then
		        f = GetFolderItem("/home/piotr/Dokumenty/prj/PaleopathologyGallery/data/" + strFolder + "/" + strFile, FolderItem.PathTypeAbsolute)
		      #ElseIf TargetMacOS Then
		        f = GetFolderItem("Macintosh HD:Users:piotrjaskulski:Documents:PaleopathologyGallery:data:" + strFolder + ":" + strFile, FolderItem.PathTypeAbsolute)
		      #EndIf
		    #Else
		      f = GetFolderItem("").Child("data").Child(strFolder).Child(strFile)
		    #EndIf
		    
		  ElseIf strType = "IMAGE" And strFile = "" Then
		    
		    #If DebugBuild Then
		      #If TargetWin32 Then
		        f = GetFolderItem("c:\PaleopathologyGallery\data\" + strFolder + "\",FolderItem.PathTypeAbsolute)
		      #Elseif TargetLinux Then
		        f = GetFolderItem("/home/piotr/Dokumenty/prj/PaleopathologyGallery/data/" + strFolder + "/",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetMacOS Then
		        f = GetFolderItem("Macintosh HD:Users:piotrjaskulski:Documents:PaleopathologyGallery:data:" + strFolder + ":",FolderItem.PathTypeAbsolute)
		      #EndIf
		    #Else
		      f = GetFolderItem("").Child("data").Child(strFolder)
		    #EndIf
		    
		  Elseif strType = "DB" Then
		    
		    #If DebugBuild Then
		      #If TargetWin32 Then
		        f = GetFolderItem("c:\PaleopathologyGallery\data\paleopathology.db",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetLinux Then
		        f = GetFolderItem("/home/piotr/Dokumenty/prj/PaleopathologyGallery/data/paleopathology.db",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetMacOS Then
		        f = GetFolderItem("Macintosh HD:Users:piotrjaskulski:Documents:PaleopathologyGallery:data:paleopathology.db",FolderItem.PathTypeAbsolute)
		      #EndIf
		    #Else
		      f = GetFolderItem("").Child("data").Child("paleopathology.db")
		    #EndIf
		    
		  Elseif strType = "BACKUP" Then
		    
		    #If DebugBuild Then
		      #If TargetWin32 Then
		        f = GetFolderItem("c:\PaleopathologyGallery\data\paleopathology.backup",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetLinux Then
		        f = GetFolderItem("/home/piotr/Dokumenty/prj/PaleopathologyGallery/data/paleopathology.backup",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetMacOS Then
		        f = GetFolderItem("Macintosh HD:Users:piotrjaskulski:Documents:PaleopathologyGallery:data:paleopathology.backup",FolderItem.PathTypeAbsolute)
		      #EndIf
		    #Else
		      f = GetFolderItem("").Child("data").Child("paleopathology.backup")
		    #EndIf
		    
		  Elseif strType = "LOGIN" Then
		    
		    #If DebugBuild Then
		      #If TargetWin32 Then
		        f = GetFolderItem("c:\PaleopathologyGallery\data\login.txt",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetLinux Then
		        f = GetFolderItem("/home/piotr/Dokumenty/prj/PaleopathologyGallery/data/login.txt",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetMacOS Then
		        f = GetFolderItem("Macintosh HD:Users:piotrjaskulski:Documents:PaleopathologyGallery:data:login.txt",FolderItem.PathTypeAbsolute)
		      #EndIf
		    #Else
		      f = GetFolderItem("").Child("data").Child("login.txt")
		    #EndIf
		    
		  Elseif strType = "HELP" Then
		    
		    #If DebugBuild Then
		      #If TargetWin32 Then
		        f = GetFolderItem("c:\PaleopathologyGallery\help\paleopathology.html",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetLinux Then
		        f = GetFolderItem("/home/piotr/Dokumenty/prj/PaleopathologyGallery/help/paleopathology.html",FolderItem.PathTypeAbsolute)
		      #ElseIf TargetMacOS Then
		        f = GetFolderItem("Macintosh HD:Users:piotrjaskulski:Documents:PaleopathologyGallery:help:paleopathology.html",FolderItem.PathTypeAbsolute)
		      #EndIf
		    #Else
		      f = GetFolderItem("").Child("help").Child("paleopathology.html")
		    #EndIf
		    
		  End If
		  
		  Return f
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ShowModally(pWindow as Window, pParentWindow as Window)
		  
		  Declare Sub EnableWindow Lib "User32" ( hwnd as Integer, enabled as Boolean )
		  Declare Function BringWindowToTop Lib "User32" ( hWnd As Integer ) As Integer
		  
		  If pParentWindow <> nil Then
		    App.mCurrentModalHandle = pParentWindow.Handle
		    EnableWindow( pParentWindow.Handle, false )
		  End if
		  
		  pWindow.ShowModal
		  
		  If pParentWindow <> Nil Then
		    EnableWindow( pParentWindow.Handle, true )
		    Dim i As Integer = BringWindowToTop ( pParentWindow.Handle )
		  End If
		  
		  App.mCurrentModalHandle = -1
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		dbPaleo As REALSQLDatabase
	#tag EndProperty

	#tag Property, Flags = &h0
		NewImages As Dictionary
	#tag EndProperty


	#tag Constant, Name = HighlightFirsColumn , Type = Boolean, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"False"
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"False"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"False"
	#tag EndConstant

	#tag Constant, Name = UseEnableWin32, Type = Boolean, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"False"
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"False"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"False"
	#tag EndConstant


	#tag Structure, Name = QueryStructure, Flags = &h0
		CAT_NO As String*512
		  SITE As String*512
		  GRAVE As String*512
		  BONE As String*512
		  SIDE As String*512
		  LOCATION As String*512
		  CHANGE_TYPE As String*512
		  PATH_TYPE As String*512
		  DIAGNOSIS As String*512
		  RunQuery As Boolean
		  ex_CAT_NO As Boolean
		  ex_SITE As Boolean
		  ex_GRAVE As Boolean
		  ex_BONE As Boolean
		  ex_SIDE As Boolean
		  ex_LOCATION As Boolean
		  ex_CHANGE_TYPE As Boolean
		  ex_PATH_TYPE As Boolean
		  ex_DIAGNOSIS As Boolean
		  SLIDE_NO As String*512
		  ex_SLIDE_NO As Boolean
		  NOTES As String*512
		ex_NOTES As Boolean
	#tag EndStructure


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
