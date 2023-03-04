#include "GUIControlEventParamInjector.ahk"
#include "BackupManager.ahk"

class PreviewHandlerManager extends GUIControlEventParamInjector
{
  Show() => this.gui.Show()
  Hide() => this.gui.Hide()

  static registryKeyNameFormat := "HKEY_CLASSES_ROOT\.{1}\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}"
      ,  noPreviewHandlerText := "None"

  __New()
  {
    static guiControlWidth := 320,
           guiButton := { w: 100, h: 24, sideMargin: 14 }

    this.__UpdatePreviewHandlersFromRegistry()

    ; Init backup formats
    this.chosenBackupFormats := Map(
        BackupManager.BackupFormat.RuntimeMemory, {},
        BackupManager.BackupFormat.RegistryFile, {
            keyNameFormat: PreviewHandlerManager.registryKeyNameFormat,
            fileNameFormat: "registryBackup_{1}.reg", 
            hourTimeFormat: RegistryFileBackup.HourTimeFormat.12,
            guidToNameCallback: (guid) => this.previewHandlers[guid]
        })
    this.chosenBackupFormats.Default := BackupManager.BackupFormat.RuntimeMemory
    this.backupManager := BackupManager(this.chosenBackupFormats)

    ; Create GUI
    this.gui := Gui(unset, "Preview Handler Manager")

    this.gui.Add("text"         , , "File extension")
    this.gui.Add("edit"         , "vGUIFileExtension w" guiControlWidth)
        .OnEvent("change", (*) => this.OnChangeExtension(this.gui["GUIFileExtension"].Value))
    this.gui.Add("text"         , , "Current preview handler associated with the extension")
    this.gui.Add("edit"         , "vGUICurrentPreviewHandler readonly w" guiControlWidth, 
        PreviewHandlerManager.noPreviewHandlerText)
    this.gui.Add("text"         , , "Choose a new preview handler for the extension")

    previewHandlerNames := this.__HumanReadablePreviewHandlerNames
    previewHandlerNames.InsertAt(1, PreviewHandlerManager.noPreviewHandlerText)
    this.gui.Add("dropdownlist" , "vGUINewPreviewHandler w" guiControlWidth " choose1", 
        previewHandlerNames).OnEvent("change", (*) => 
            this.OnChangePreviewHandlerSelection(this.gui["GUIFileExtension"].Value))

    this.gui.Add("checkbox"     , "vGUIBackup checked y+20 section", "Use backups")
        .OnEvent("click", (*) => this.OnToggleCreateBackups(this.gui["GUIFileExtension"].Value))
    this.gui.Add("button"       , "vGUIRestore yp-5 w" guiButton.w " h" guiButton.h 
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

    _OnChangeExtension(fileExtension)
    {
      ; Validation rule: optional dot followed by some word char(s)
      isValidExtension := fileExtension ~= "^\w+$"

      currentPreviewHandlerGuid := regread(
          Format(PreviewHandlerManager.registryKeyNameFormat, fileExtension), , 0)

      ; Display current handler
      this.gui["GUICurrentPreviewHandler"].Value := currentPreviewHandlerGuid 
          ? this.previewHandlers[currentPreviewHandlerGuid] 
          : PreviewHandlerManager.noPreviewHandlerText

      ; Enable only when a valid extension and its currently active 
      ; preview handler matches the one selected in the dropdownlist
      this.gui["GUIBind"].Enabled := isValidExtension 
          && currentPreviewHandlerGuid !== this.__SelectedPreviewHandlerGuid

      ; Enable only when the backup feature is active and the user has typed an already
      ; backed-up extension whose currently set preview handler differs from that of the backup.
      backupFormat := this.chosenBackupFormats.Default
      this.gui["GUIRestore"].Enabled := this.gui["GUIBackup"].Value
          && this.backupManager.IsAlreadyCreatedForSession(fileExtension, backupFormat)
          && this.backupManager.Retrieve(fileExtension, backupFormat) != currentPreviewHandlerGuid
    }

    OnChangePreviewHandlerSelection(fileExtension)
    {
      isValidExtension := fileExtension ~= "^\w+$"
      
      currentPreviewHandlerGuid := regread(
          Format(PreviewHandlerManager.registryKeyNameFormat, fileExtension), , 0)

      ; Enable only if valid extension and currently active preview 
      ; handler does not match the one selected in the dropdownlist
      this.gui["GUIBind"].Enabled := isValidExtension && 
          currentPreviewHandlerGuid !== this.__SelectedPreviewHandlerGuid
    }

    _OnSetPreviewHandler(fileExtension)
    {
      registryKeyName := Format(PreviewHandlerManager.registryKeyNameFormat, fileExtension)

      currentPreviewHandlerGuid := regread(
          Format(PreviewHandlerManager.registryKeyNameFormat, fileExtension), , 0)

      ; Create backup if not already created
      if (this.gui["GUIBackup"].Value)
      {
        this.__CreateBackup(fileExtension, currentPreviewHandlerGuid)
      }

      selectedPreviewHandlerGuid := this.__SelectedPreviewHandlerGuid

      if (selectedPreviewHandlerGuid)
      {
        ; Other than "None" selected
        regwrite selectedPreviewHandlerGuid, "REG_SZ", registryKeyName
        newlySetHandlerName := this.previewHandlers[selectedPreviewHandlerGuid]
      }
      else if (currentPreviewHandlerGuid)
      {
        ; "None" selected
        regdeletekey registryKeyName
        newlySetHandlerName := PreviewHandlerManager.noPreviewHandlerText
      }

      this.gui["GUICurrentPreviewHandler"].Value := newlySetHandlerName

      ; Enable only when the backup feature is active and the preview
      ; handler selected in the dropdownlist (i.e. the one that was
      ; just set as the active handler) differs from that in the backup
      this.gui["GUIRestore"].Enabled := this.gui["GUIBackup"].Value 
          && (this.backupManager.Retrieve(fileExtension, this.chosenBackupFormats.Default) 
              !== selectedPreviewHandlerGuid)
      
      this.gui["GUIBind"].Enabled := false

      msgbox
      (
        "Preview handler for the extension (." fileExtension ") "
        "has been set to '" newlySetHandlerName "'."
      )
    }

    _OnRestorePreviewHandler(fileExtension)
    {
      originalPreviewHandlerGuid := this.__RestoreBackup(fileExtension)
      
      this.gui["GUICurrentPreviewHandler"].Value := originalPreviewHandlerGuid 
          ? this.previewHandlers[originalPreviewHandlerGuid]
          : PreviewHandlerManager.noPreviewHandlerText

      ; Enable only when the restored handler differs from that selected in the dropdownlist
      this.gui["GUIBind"].Enabled := originalPreviewHandlerGuid !== this.__SelectedPreviewHandlerGuid

      this.gui["GUIRestore"].Enabled := false

      msgbox (originalPreviewHandlerGuid ?
      (c
        ; Backup was other than "None"
        "Original preview handler for the extension (." fileExtension ") "
        "has been restored back to initial '" this.previewHandlers[originalPreviewHandlerGuid] "'."
      ) :
      (c
        ; Backup was "None"
        "Preview handler for the extension (." fileExtension ") has been removed."
      ))
    }

    _OnToggleCreateBackups(fileExtension)
    {
      currentPreviewHandlerGuid := regread(
          Format(PreviewHandlerManager.registryKeyNameFormat, fileExtension), , 0)

      backupFormat := this.chosenBackupFormats.Default

      ; Enable only when the backup feature is active, a backup has been created earlier
      ; and the currently active preview handler differs from that of the backup
      this.gui["GUIRestore"].Enabled := this.gui["GUIBackup"].Value
          && this.backupManager.IsAlreadyCreatedForSession(fileExtension, backupFormat)
          && (this.backupManager.Retrieve(fileExtension, backupFormat) 
              !== currentPreviewHandlerGuid)
    }

  ; ============================================================
  ; Private methods
  ; ============================================================

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
      registryKeyName := Format(PreviewHandlerManager.registryKeyNameFormat, fileExtension)

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

  ; ============================================================
  ; Private dynamic properties
  ; ============================================================

    ; Preview handler GUID matching the currently selected dropdown item
    __SelectedPreviewHandlerGuid
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
}