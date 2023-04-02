#include "FileExtensionParamSanitizer.ahk"
#include "BackupManager.ahk"

class PreviewHandlerManager extends FileExtensionParamSanitizer
{
  Show() => this.gui.Show()
  Hide() => this.gui.Hide()

  static REGISTRY_KEYNAME_FORMAT := "HKEY_CLASSES_ROOT\.{1}\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}"
      ,  PREVIEW_HANDLER_TEXT := { NONE: "None", UNKNOWN: "Unknown: {1}" }

  __New()
  {
    static guiControlWidth := 320,
          guiButton := { w: 100, h: 24, sideMargin: 14 }

    if (this.HasOwnProp("gui"))
    {
      return
    }
    
    super.__New()

    this.__UpdatePreviewHandlersFromRegistry()

    ; Init backup formats
    this.chosenBackupFormats := Map(
        BackupManager.BackupFormat.RuntimeMemory, {},
        BackupManager.BackupFormat.RegistryFile, {
            keyNameFormat: this.__GetPreviewHandlerRegistryLocation,
            fileNameFormat: "registryBackup_{1}.reg", 
            hourTimeFormat: RegistryFileBackup.HourTimeFormat.12,
            guidToNameCallback: (guid) => this.previewHandlers.Has(guid) 
                ? this.previewHandlers[guid] 
                : Format(PreviewHandlerManager.PREVIEW_HANDLER_TEXT.% guid != 0 ? "UNKNOWN" : "NONE" %, guid)
        })
    this.chosenBackupFormats.Default := BackupManager.BackupFormat.RuntimeMemory
    this.backupManager := BackupManager(this.chosenBackupFormats)

    ; Create GUI
    this.gui := Gui(unset, "Preview Handler Manager")
    
    this.gui.Add("text"         , "y+10 x+10", "File extension")
    this.gui.Add("edit"         , "vGUIFileExtension limit w" guiControlWidth)
        .OnEvent("change", (*) => this.OnChangeExtension(this.gui["GUIFileExtension"].Value))
    this.gui.Add("groupbox"     , "vGUIGroupBox xp-10 y+10 w" (guiControlWidth + 20) " h154", "")
    this.gui.Add("text"         , "xp+10 yp+20", "Current preview handler associated with the extension")
    this.gui.Add("edit"         , "vGUICurrentPreviewHandler readonly "
        "w" guiControlWidth, PreviewHandlerManager.PREVIEW_HANDLER_TEXT.NONE)
    this.gui.Add("text"         , , "Choose a new preview handler for the extension")

    previewHandlerNames := this.__HumanReadablePreviewHandlerNames
    previewHandlerNames.InsertAt(1, PreviewHandlerManager.PREVIEW_HANDLER_TEXT.NONE)
    this.gui.Add("dropdownlist" , "vGUINewPreviewHandler w" guiControlWidth " choose1", 
        previewHandlerNames).OnEvent("change", (*) => 
            this.OnChangePreviewHandlerSelection(this.gui["GUIFileExtension"].Value))

    ; this.gui.Add("checkbox"     , "vGUIBackup checked y+20 section", "Use backups")
        ; .OnEvent("click", (*) => this.OnToggleCreateBackups(this.gui["GUIFileExtension"].Value))
    this.gui.Add("button"       , "vGUIRestore y+15 w" guiButton.w " h" guiButton.h 
        " disabled", "Restore original")
        .OnEvent("click", (*) => this.OnRestorePreviewHandler(this.gui["GUIFileExtension"].Value))
    this.gui.Add("button"       , "vGUIBind default yp w" guiButton.w " h" guiButton.h " "
        " disabled", "Bind handler")
        .OnEvent("click", (*) => this.OnSetPreviewHandler(this.gui["GUIFileExtension"].Value))

    this.gui.Show("hide")

    ; Reposition buttons to align to right edge
    this.gui["GUINewPreviewHandler"].GetPos(&dropdownX, , &dropdownW)
    this.gui["GUIBind"].GetPos(, , &buttonWidth)
    this.gui["GUIBind"].Move(dropdownX + dropdownW - buttonWidth)
    this.gui["GUIRestore"].Move(dropdownX + dropdownW - buttonWidth * 2 - guiButton.sideMargin)

    this.gui.OnEvent("close", (*) => exitapp())
  }

  ; ============================================================
  ; GUI methods
  ; ============================================================

    OnChangeExtension(fileExtension)
    {
      static timerInfo := { triggerMs: 700 }

      this.gui["GUIGroupbox"].Text := fileExtension !== "" ? "." fileExtension : ""

      ; Validation rule: optional dot followed by some word char(s)
      isValidExtension := fileExtension ~= "^\w+$"

      currentRegistryPreviewHandler := this.__RegistryPreviewHandler[fileExtension]

      ; Set background color to display during simulated value retrieval delay
      this.gui["GUICurrentPreviewHandler"].Opt("+" "background0xdddddd redraw")
      
      ; Enable only when a valid extension and its currently active 
      ; preview handler doesn't match the one selected in the dropdownlist
      this.gui["GUIBind"].Enabled := isValidExtension 
          && currentRegistryPreviewHandler.guid !== this.__DropdownPreviewHandlerGuid

      ; Enable only when the backup feature is active and the user has typed an already
      ; backed-up extension whose currently set preview handler differs from that of the backup.
      backupFormat := this.chosenBackupFormats.Default
      this.gui["GUIRestore"].Enabled := true ;this.gui["GUIBackup"].Value
          && this.backupManager.IsAlreadyCreatedForSession(fileExtension, backupFormat)
          && this.backupManager.Retrieve(fileExtension, backupFormat) 
              != currentRegistryPreviewHandler.guid
      
      timerInfo.countDownTickCount := a_tickcount
      settimer(this.__DelayedChangeExtension.Bind(this, 
          currentRegistryPreviewHandler, timerInfo), -timerInfo.triggerMs)
    }

    OnChangePreviewHandlerSelection(fileExtension)
    {
      isValidExtension := fileExtension ~= "^\w+$"
      
      currentRegistryPreviewHandler := this.__RegistryPreviewHandler[fileExtension]

      ; Enable only if valid extension and currently active preview 
      ; handler does not match the one selected in the dropdownlist
      this.gui["GUIBind"].Enabled := isValidExtension && 
          currentRegistryPreviewHandler.guid !== this.__DropdownPreviewHandlerGuid
    }

    OnSetPreviewHandler(fileExtension)
    {
      registryKeyName := this.__GetPreviewHandlerRegistryLocation(fileExtension)

      currentRegistryPreviewHandler := this.__RegistryPreviewHandler[fileExtension]

      ; Create backup if not already created
      ; if (this.gui["GUIBackup"].Value)
      {
        this.__CreateBackup(fileExtension, currentRegistryPreviewHandler.guid)
      }

      selectedPreviewHandlerGuid := this.__DropdownPreviewHandlerGuid

      if (selectedPreviewHandlerGuid)
      {
        ; Other than "None" selected
        regwrite selectedPreviewHandlerGuid, "REG_SZ", registryKeyName
        newlySetHandlerName := this.previewHandlers[selectedPreviewHandlerGuid]
      }
      else if (currentRegistryPreviewHandler.guid)
      {
        ; "None" selected and a preview handler currently assigned to the extension
        regdeletekey registryKeyName
        newlySetHandlerName := PreviewHandlerManager.PREVIEW_HANDLER_TEXT.NONE
      }

      this.gui["GUICurrentPreviewHandler"].Value := newlySetHandlerName

      ; Enable only when the backup feature is active and the preview
      ; handler selected in the dropdownlist (i.e. the one that was
      ; just set as the active handler) differs from that in the backup
      this.gui["GUIRestore"].Enabled := true ;this.gui["GUIBackup"].Value 
          && (this.backupManager.Retrieve(fileExtension, this.chosenBackupFormats.Default) 
              !== selectedPreviewHandlerGuid)
      
      this.gui["GUIBind"].Enabled := false

      msgbox
      (
        "Preview handler for the extension (." fileExtension ") "
        "has been set to '" newlySetHandlerName "'."
      )
    }

    OnRestorePreviewHandler(fileExtension)
    {
      originalPreviewHandlerGuid := this.__RestoreBackup(fileExtension)
      
      if (this.previewHandlers.Has(originalPreviewHandlerGuid))
      {
        originalPreviewHandlerName := this.previewHandlers[originalPreviewHandlerGuid]
      }
      else
      {
        originalPreviewHandlerName := Format(PreviewHandlerManager
            .PREVIEW_HANDLER_TEXT.% originalPreviewHandlerGuid ? "UNKNOWN" : "NONE" %, 
            originalPreviewHandlerGuid)
      }

      this.gui["GUICurrentPreviewHandler"].Value := originalPreviewHandlerName

      ; Enable only when the restored handler differs from that selected in the dropdownlist
      this.gui["GUIBind"].Enabled := originalPreviewHandlerGuid 
          !== this.__DropdownPreviewHandlerGuid

      this.gui["GUIRestore"].Enabled := false

      msgbox (originalPreviewHandlerGuid ?
      (c
        ; Backup was other than "None"
        "Original preview handler for the extension (." fileExtension ") "
        "has been restored back to initial '" originalPreviewHandlerName "'."
      ) :
      (c
        ; Backup was "None"
        "Preview handler for the extension (." fileExtension ") has been removed."
      ))
    }

    OnToggleCreateBackups(fileExtension)
    {
      currentRegistryPreviewHandler := this.__RegistryPreviewHandler[fileExtension]

      backupFormat := this.chosenBackupFormats.Default

      ; Enable only when the backup feature is active, a backup has been created earlier
      ; and the currently active preview handler differs from that of the backup
      this.gui["GUIRestore"].Enabled := true ;this.gui["GUIBackup"].Value
          && this.backupManager.IsAlreadyCreatedForSession(fileExtension, backupFormat)
          && (this.backupManager.Retrieve(fileExtension, backupFormat) 
              !== currentRegistryPreviewHandler.guid)
    }
  
  ; ============================================================
  ; Private methods
  ; ============================================================

    ; Restore original bkgnd color to the Edit control and display the current preview handler
    __DelayedChangeExtension(currentRegistryPreviewHandler, timerInfo)
    {
      if (a_tickcount - timerInfo.countDownTickCount < timerInfo.triggerMs)
      {
        ; This SetTimer invocation pushed back by a later instance
        return
      }

      ; Restore default background color
      this.gui["GUICurrentPreviewHandler"].Opt("+background redraw")

      ; Display current handler
      if (currentRegistryPreviewHandler.isRecognized)
      {
        this.gui["GUICurrentPreviewHandler"].Value := currentRegistryPreviewHandler.guid
            ? this.previewHandlers[currentRegistryPreviewHandler.guid]
            : PreviewHandlerManager.PREVIEW_HANDLER_TEXT.NONE
      }
      else
      {
        this.gui["GUICurrentPreviewHandler"].Value 
            := Format(PreviewHandlerManager.PREVIEW_HANDLER_TEXT.UNKNOWN, 
                currentRegistryPreviewHandler.guid)
      }
    }

    ; Backup the extension's original preview handler
    __CreateBackup(fileExtension, valueToBackup)
    {
      backupFormatNames := []
      for backupFormatName in this.chosenBackupFormats
      {
        if (!this.backupManager.IsAlreadyCreatedForSession(fileExtension, backupFormatName))
        {
          backupFormatNames.Push(backupFormatName)
        }
      }

      if (backupFormatNames.Length)
      {
        this.backupManager.Create(fileExtension, valueToBackup, backupFormatNames*)
      }
    }

    ; Restore extension's original handler
    __RestoreBackup(fileExtension)
    {
      registryKeyName := this.__GetPreviewHandlerRegistryLocation(fileExtension)

      backupGuid := this.backupManager.Retrieve(fileExtension, this.chosenBackupFormats.Default)

      if (backupGuid)
      {
        ; Extension initially had a preview handler associated with it

        regwrite backupGuid, "REG_SZ", registryKeyName
      }
      else
      {
        ; Extension initially had no existing preview handler

        try
        {
          regdeletekey registryKeyName
        }
        catch Error as e
        {
          outputdebug Format("Unexpected app state: Registry key '{1}' could't be deleted: {2}", 
              registryKeyName, e.Message)
        }

        ; Remove also "ShellEx" key if empty
        shellExKey := regexreplace(registryKeyName, "\\[^\\]+$")
        loop reg shellExKey, "KV"
        {
          break
        }
        else
        {
          try
          {
            regdeletekey shellExKey
          }
          catch Error as e
          {
            outputdebug Format("Unexpected app state: Registry key '{1}' could't be deleted: {2}",
                shellExKey, e.Message)
          }
        }
      }

      return backupGuid
    }

    __UpdatePreviewHandlersFromRegistry()
    {
      static previewHandlerRegistryPath 
          := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers"

      this.previewHandlers := Map()

      ; Reg path appears to contain only values, so skipping checking for cascaded keys here
      loop reg previewHandlerRegistryPath
      {
        this.previewHandlers[a_loopregname] := regread()
      }
    }

    __GetPreviewHandlerRegistryLocation(fileExtension)
    {
      static WINAPI :=
      {
        ERROR_SUCCESS: 0,
        HKEY_CLASSES_ROOT: 0x80000000,
        KEY_READ: 0x20019
      }, hKey := 0
      , validPreviewHandlerValuePattern := "i)^\{[0-9A-F]{8}-([0-9A-F]{4}-){3}[0-9A-F]{12}\}$"

      primaryLocationFoundButInvalid := secondaryLocationFoundButInvalid := false

      ; HKCR\.reg\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
      secondaryRegistryLocation := Format(
          PreviewHandlerManager.REGISTRY_KEYNAME_FORMAT, fileExtension)

      ; HKCR\.reg
      delegatePrimaryKeyLocation := regexreplace(secondaryRegistryLocation, "\\ShellEx.*$")

      ; regfile
      delegatePrimaryKeyName := regread(delegatePrimaryKeyLocation, , 0)

      if (delegatePrimaryKeyName)
      {
        ; HKCR\regfile\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
        primaryRegistryLocation := Format(strreplace(
            PreviewHandlerManager.REGISTRY_KEYNAME_FORMAT, "."), delegatePrimaryKeyName)
        
        if (keyExists := WINAPI.ERROR_SUCCESS == DllCall("advapi32\RegOpenKeyEx", 
            "ptr", WINAPI.HKEY_CLASSES_ROOT, 
            "str", RegExReplace(primaryRegistryLocation, "^[A-Z_]+\\"),
            "int", 0, 
            "int", WINAPI.KEY_READ := 0x20019,
            "ptr*", hKey))
        {
          defaultValue := regread(primaryRegistryLocation, , "")
          
          if (regexmatch(defaultValue, validPreviewHandlerValuePattern))
          {
            return primaryRegistryLocation
          }

          primaryLocationFoundButInvalid := true
        }
      }
      ; HKCR\.reg\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
      if (keyExists := WINAPI.ERROR_SUCCESS == DllCall("advapi32\RegOpenKeyEx", 
          "ptr", WINAPI.HKEY_CLASSES_ROOT, 
          "str", RegExReplace(secondaryRegistryLocation, "^[A-Z_]+\\"),
          "int", 0, 
          "int", WINAPI.KEY_READ := 0x20019,
          "ptr*", hKey))
      {
        defaultValue := regread(secondaryRegistryLocation, , "")

        if (regexmatch(defaultValue, validPreviewHandlerValuePattern))
        {
          return secondaryRegistryLocation
        }

        secondaryLocationFoundButInvalid := true
      }

      if (primaryLocationFoundButInvalid)
      {
        return primaryRegistryLocation
      }

      if (secondaryLocationFoundButInvalid)
      {
        return secondaryRegistryLocation
      }

      return 0
    }
  
  ; ============================================================
  ; Private dynamic properties
  ; ============================================================

    ; Preview handler GUID matching the currently selected dropdown item
    __DropdownPreviewHandlerGuid
    {
      get
      {
        for guid, humanReadableName in this.previewHandlers
        {
          ; Account for the "None" item at the beginning of the dropdownlist
          if (a_index == this.gui["GUINewPreviewHandler"].Value - 1)
          {
            return guid
          }
        }

        return 0
      }
    }

    ; Preview handler GUID of the given extension in the registry
    __RegistryPreviewHandler[fileExtension]
    {
      get {
        result :=
        {
          guid: 0
        }

        registryKey := this.__GetPreviewHandlerRegistryLocation(fileExtension)
        
        if (registryKey)
        {
          extensionGuid := regread(registryKey)

          if (extensionGuid != "")
          {
            result.guid := extensionGuid
          }
        }
        
        ; A preview handler may be set up in the primary location
        ; HKCR\{.[fileExtension] (Default) value}\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
        ; or in the secondary location
        ; HKCR\.[fileExtension]\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
        ; but not available in
        ; HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers
        result.isRecognized := result.guid == 0 || this.previewHandlers.Has(result.guid)

        return result
      }
    }

    ; From map values into array
    __HumanReadablePreviewHandlerNames
    {
      get
      {
        result := []

        for , value in this.previewHandlers
        {
          result.Push(value)
        }

        return result
      }
    }

    __Throw(err)
    {
      throw err
    }

    static __GlobalExceptionHandler(error, mode)
    {
      msgbox "An unhandled exception occurred."
    }
}