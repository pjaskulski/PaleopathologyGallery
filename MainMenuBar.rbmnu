#tag Menu
Begin Menu MainMenuBar
   Begin MenuItem FileMenu
      SpecialMenu = 0
      Text = "&Records"
      Index = -2147483648
      AutoEnable = True
      Begin MenuItem NewItem
         SpecialMenu = 0
         Text = "New"
         Index = -2147483648
         ShortcutKey = "N"
         Shortcut = "Cmd+N"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem EditItem
         SpecialMenu = 0
         Text = "Edit"
         Index = -2147483648
         ShortcutKey = "E"
         Shortcut = "Cmd+E"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem DeleteItem
         SpecialMenu = 0
         Text = "Delete"
         Index = -2147483648
         ShortcutKey = "D"
         Shortcut = "Cmd+D"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem UntitledSeparator2
         SpecialMenu = 0
         Text = "-"
         Index = -2147483648
         AutoEnable = True
      End
      Begin MenuItem ImportItem
         SpecialMenu = 0
         Text = "Import"
         Index = -2147483648
         ShortcutKey = "O"
         Shortcut = "Cmd+O"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem UntitledSeparator3
         SpecialMenu = 0
         Text = "-"
         Index = -2147483648
         AutoEnable = True
      End
      Begin MenuItem DBSynchronization
         SpecialMenu = 0
         Text = "Synchronization"
         Index = -2147483648
         ShortcutKey = "R"
         Shortcut = "Cmd+R"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem UntitledSeparator4
         SpecialMenu = 0
         Text = "-"
         Index = -2147483648
         AutoEnable = True
      End
      Begin MenuItem FindItem
         SpecialMenu = 0
         Text = "F&ind"
         Index = -2147483648
         ShortcutKey = "I"
         Shortcut = "Cmd+I"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem FindNextItem
         SpecialMenu = 0
         Text = "Find &next"
         Index = -2147483648
         ShortcutKey = "F3"
         Shortcut = "F3"
         AutoEnable = True
      End
      Begin MenuItem UntitledSeparator1
         SpecialMenu = 0
         Text = "-"
         Index = -2147483648
         AutoEnable = True
      End
      Begin MenuItem FilterItem
         SpecialMenu = 0
         Text = "&Filter"
         Index = -2147483648
         ShortcutKey = "F"
         Shortcut = "Cmd+F"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem ShowAllItem
         SpecialMenu = 0
         Text = "Show all records"
         Index = -2147483648
         ShortcutKey = "G"
         Shortcut = "Cmd+G"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem UntitledSeparator0
         SpecialMenu = 0
         Text = "-"
         Index = -2147483648
         AutoEnable = True
      End
      Begin MenuItem ZoomItem
         SpecialMenu = 0
         Text = "Zoom"
         Index = -2147483648
         ShortcutKey = "M"
         Shortcut = "Cmd+M"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem CopyFileToItem
         SpecialMenu = 0
         Text = "Copy file to..."
         Index = -2147483648
         ShortcutKey = "P"
         Shortcut = "Cmd+P"
         MenuModifier = True
         AutoEnable = True
      End
      Begin MenuItem DbExport
         SpecialMenu = 0
         Text = "Export..."
         Index = -2147483648
         AutoEnable = True
      End
      Begin MenuItem UntitledSeparator
         SpecialMenu = 0
         Text = "-"
         Index = -2147483648
         AutoEnable = True
      End
      Begin QuitMenuItem FileQuit
         SpecialMenu = 0
         Text = "#App.kFileQuit"
         Index = -2147483648
         ShortcutKey = "#App.kFileQuitShortcut"
         Shortcut = "#App.kFileQuitShortcut"
         AutoEnable = True
      End
   End
   Begin MenuItem HelpMenu
      SpecialMenu = 0
      Text = "Help"
      Index = -2147483648
      AutoEnable = True
      Begin AppleMenuItem HelpAbout
         SpecialMenu = 0
         Text = "#App.kHelpAbout"
         Index = -2147483648
         AutoEnable = True
      End
      Begin MenuItem HelpHelp
         SpecialMenu = 0
         Text = "Help"
         Index = -2147483648
         ShortcutKey = "F1"
         Shortcut = "F1"
         AutoEnable = True
      End
   End
End
#tag EndMenu
