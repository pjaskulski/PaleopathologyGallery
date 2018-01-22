#tag Class
Protected Class DoubleClickImage
Inherits ImageWell
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  return true
		End Function
	#tag EndEvent

	#tag Event
		Sub MouseUp(X As Integer, Y As Integer)
		  dim doubleClickTime, currentClickTicks as Integer
		  
		  #if TargetMacOS then
		    #if TargetCarbon then
		      Declare Function GetDblTime Lib "Carbon" () as Integer
		    #else
		      Declare Function GetDblTime Lib "InterfaceLib" () as Integer Inline68K("2EB802F0")
		    #endif
		    doubleClickTime = GetDblTime()
		  #endif
		  
		  #if TargetWin32 then
		    Declare Function GetDoubleClickTime Lib "User32.DLL" () as Integer
		    doubleClickTime = GetDoubleClickTime()
		  #endif
		  
		  #if TargetLinux then
		    Declare Function gtk_settings_get_default lib "libgtk-x11-2.0.so" as Ptr
		    Declare Sub g_object_get lib "libgtk-x11-2.0.so" (Obj as Ptr, first_property_name as CString, byref doubleClicktime as Integer, Null as Integer)
		    
		    dim gtkSettings as MemoryBlock
		    
		    gtkSettings = gtk_settings_get_default()
		    
		    g_object_get(gtkSettings,"gtk-double-click-time",doubleClickTime, 0)
		    // DoubleClickTime now holds the number of milliseconds
		    doubleClickTime = doubleClickTime / 1000.0 * 60
		  #endif
		  
		  currentClickTicks = ticks
		  //if the two clicks happened close enough together in time
		  if (currentClickTicks - lastClickTicks) <= doubleClickTime then
		    //if the two clicks occured close enough together in space
		    if abs(X - lastClickX) <= 5 and abs(Y - LastClickY) <= 5 then
		      DoubleClick //a double click has occured so call the event
		    end if
		  end if
		  lastClickTicks = currentClickTicks
		  lastClickX = X
		  lastClickY = Y
		End Sub
	#tag EndEvent


	#tag Hook, Flags = &h0
		Event DoubleClick()
	#tag EndHook


	#tag Property, Flags = &h1
		Protected lastClickTicks As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected lastClickX As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected lastClickY As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			Type="Integer"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Width"
			Visible=true
			Group="Position"
			InitialValue="32"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Height"
			Visible=true
			Group="Position"
			InitialValue="32"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockLeft"
			Visible=true
			Group="Position"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockTop"
			Visible=true
			Group="Position"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockRight"
			Visible=true
			Group="Position"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockBottom"
			Visible=true
			Group="Position"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabPanelIndex"
			Group="Position"
			InitialValue="0"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabIndex"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabStop"
			Visible=true
			Group="Position"
			InitialValue="True"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="InitialParent"
			Group="Position"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Visible"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HelpTag"
			Visible=true
			Group="Appearance"
			EditorType="MultiLineEditor"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AutoDeactivate"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Enabled"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Image"
			Visible=true
			Group="Appearance"
			EditorType="Picture"
			InheritedFrom="ImageWell"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
