/************************************************************************
 * @description Manage Windows Explorer preview panel file extension handlers
 * @script Preview Handler Manager
 * @file PreviewHandlerManager.ahk
 * @author TJay🐦 <sales@rpamania.net> Preview Handler Manager
 * @author RaptorX <graptorx@gmail.com> original ScriptObj before modifications
 * @date 2023-04-13
 * @version 1.0.0
 ***********************************************************************/
script.version := "1.0.0"

/*
  ** Notes by TJay **
  Mar 2, 2023:
    - Add/modify a preview handler per file extension
    - Allow creating a backup of the original handler and restoring it
      * Create an individual backup record for each modified extension
        > However, all extension handler backups will be appended into a single registry
          file if the RegistryFile backup format is active (hardcoded on by default)
      * Create a runtime memory backup
      * Create a registry file backup (.reg file, saved to the script folder)
    - Supports preview handlers specified in the registry branch 
      HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers
      * Creates new handler entries in 
        HKCR\[.ext]\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
      * Detect + allow backing up and restoring unknown preview handlers specified
        for a file extension
    - "Use backups" checkbox
      * New backups won't not created while unticked
      * If unticked after a backup has been created, restoring the backup will be 
        disabled until ticked back on
      * Ticking/unticking while backups have already been created during a 
        session may possibly have unexpected consequences (not thoroughly tested)
  
  Mar 11, 2023:
    - Fixed a crash caused by a preview handler having been specified for a file 
      extension in the registry while being absent in 
      HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PreviewHandlers
  
  Mar 29, 2023:
    - Removed "Use backups" checkbox → the backup feature is now fixed to always remain on
    - Modified the UI to emphasize automatic retrieval of handler info for a current extension
      * Added a groupbox with the typed extension updating in the title
      * Changed UI control color with a timer to signal update on changing of an extension
  
  Apr 5, 2023:
    - Implemented handling of the alternate (=primary) preview handler location in the registry,
      HKCR\{HKCR\.ext\(Default) value}\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f},
      such as HKCR\regfile\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f} for .reg files
    - Improve validation and handling of file extension input by a user + apply custom red 
      background color to the file extension Edit field to indicate invalid input
  
  Apr 12, 2023:
    - Fix .htm[l] extension handling where the primary registry location is write-protected
    - Add modified version of RaptorX's ScriptObj to enable auto-update
  
  TODO:
    - Fix UI DPI scaling
    - Allow defining a catch-all handler: HKCR\*\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
    - Display the name of the original (backed-up) preview handler of an extension? 
*/

#requires autohotkey v2.0
#singleinstance ignore
#include <FileExtensionParamSanitizer>
#include <ScriptObj>
#include <BackupManager>
#notrayicon

if (!A_IsAdmin) { ;http://ahkscript.org/docs/Variables.htm#IsAdmin
  Run "*RunAs `"" A_ScriptFullPath "`""  ; Requires v1.0.92.01+
  ExitApp
}

if (a_iscompiled)
{
  OnError PreviewHandlerManager.__GlobalExceptionHandler.Bind(PreviewHandlerManager)
}
else
{
  a_iconhidden := false
}

PreviewHandlerManager().Show()



class PreviewHandlerManager extends FileExtensionParamSanitizer
{
  Show() => this.gui.Show()
  Hide() => this.gui.Hide()

  static REGISTRY_KEYNAME_FORMAT := "HKEY_CLASSES_ROOT\.{1}\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}"
      ,  PREVIEW_HANDLER_TEXT := { NONE: "None", UNKNOWN: "Unknown: '{1}'" }
      ,  validFileExtensionPattern := "^[^\\/:*?`"<>|]+$"
  
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
            fileNameFormat: "registryBackup_{1}.reg", 
            hourTimeFormat: RegistryFileBackup.HourTimeFormat.12,
            guidToNameCallback: (guid) => this.previewHandlers.Has(guid) 
                ? this.previewHandlers[guid] 
                : Format(PreviewHandlerManager.PREVIEW_HANDLER_TEXT.% guid != 0 
                    ? "UNKNOWN" : "NONE" %, guid)
        })
    this.chosenBackupFormats.Default := BackupManager.BackupFormat.RuntimeMemory
    this.backupManager := BackupManager(this.chosenBackupFormats)

    ; Create GUI
    this.gui := Gui(, "Preview Handler Manager")
    
    fileMenu := Menu()
    fileMenu.Add("&Check for Updates", (*) => 
        this.__MenuCheckUpdate("https://github.com/RPAMania/" 
            PreviewHandlerManager.Prototype.__Class "/releases/latest"))

    menus := MenuBar()
    menus.Add("&File", fileMenu)

    this.gui.MenuBar := menus
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
        " disabled", "Restore initial")
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

    isValidExtension := fileExtension ~= PreviewHandlerManager.validFileExtensionPattern

    if (isValidExtension || fileExtension == "")
    {
      this.gui["GUIFileExtension"].Opt("+background redraw")
    }
    else
    {
      this.gui["GUIFileExtension"].Opt("+" "background0xffaaaa redraw")
    }

    currentRegistryPreviewHandler := this.__RegistryPreviewHandler[fileExtension]
    
    if (!isValidExtension)
    {
      this.gui["GUIBind"].Enabled := false
      this.gui["GUIRestore"].Enabled := false
      return
    }

    ; Enable only when a valid extension and its currently active 
    ; preview handler doesn't match the one selected in the dropdownlist
    this.gui["GUIBind"].Enabled := isValidExtension 
        && currentRegistryPreviewHandler.guid !== this.__DropdownPreviewHandlerGuid

    ; Set background color to display during simulated value retrieval delay
    this.gui["GUICurrentPreviewHandler"].Opt("+" "background0xdddddd redraw")

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
    isValidExtension := fileExtension ~= PreviewHandlerManager.validFileExtensionPattern
    
    currentRegistryPreviewHandler := this.__RegistryPreviewHandler[fileExtension]

    ; Enable only if valid extension and currently active preview 
    ; handler does not match the one selected in the dropdownlist
    this.gui["GUIBind"].Enabled := isValidExtension && 
        currentRegistryPreviewHandler.guid !== this.__DropdownPreviewHandlerGuid
  }

  OnSetPreviewHandler(fileExtension)
  {
    registryKeyInfo := this.__GetPreviewHandlerRegistryLocation(fileExtension)

    currentRegistryPreviewHandler := this.__RegistryPreviewHandler[fileExtension]

    ; if (this.gui["GUIBackup"].Value)
    {
      ; Create backup if not already created

      ; Current detected preview handler may not always be the one to back up.
      ; For example, .htm and .html extensions in Windows 10 are treated differently 
      ; than many others by Windows Explorer in that the ordinary primary location 
      ; (in the case of .htm[l] files, HKCR\htmlfile\ShellEx) is write-protected and
      ; actually used as a secondary location inferior to HKCR\.htm[l] in terms 
      ; of precedence. This means that to back up the existing preview handler when 
      ; the ordinary secondary (== new primary) location is not set or has empty value 
      ; and the ordinary primary (== new secondary) location is write protected, we 
      ; must actually back up the value of the new primary location.
      ; If the value was empty/missing, i.e. the new primary location didn't have a 
      ; preview handler but the write-protected new secondary location potentially had 
      ; one, restoring the backup will wipe the newly set handler from the new primary 
      ; location, leaving any existing handler in the new secondary location for the
      ; Windows Explorer to fall back to, as it initially was.
      if (registryKeyInfo.primaryLocationCandidate.isWriteProtected &&
          ((!registryKeyInfo.actualLocationCandidate.isFound) ||
          regread(registryKeyInfo.actualLocationCandidate.branch, , "") == ""))
      {
        guidToBackup := 0
      }
      else
      {
        guidToBackup := currentRegistryPreviewHandler.guid
      }

      this.__CreateBackup(fileExtension, {
          guid: guidToBackup,
          registryBranch: registryKeyInfo.actualLocationCandidate.branch
      })
    }

    selectedPreviewHandlerGuid := this.__DropdownPreviewHandlerGuid

    if (selectedPreviewHandlerGuid)
    {
      ; Other than "None" selected
      ; DllCall("Reg")
      regwrite selectedPreviewHandlerGuid, "REG_SZ", 
          registryKeyInfo.actualLocationCandidate.branch
      newlySetHandlerName := this.previewHandlers[selectedPreviewHandlerGuid]
    }
    else if (currentRegistryPreviewHandler.guid)
    {
      ; "None" selected and a preview handler currently assigned to the extension
      regdeletekey registryKeyInfo.actualLocationCandidate.branch
      this.__RegRemoveShellExIfEmpty(registryKeyInfo.actualLocationCandidate.branch)

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
      "Preview handler for the extension ." fileExtension " "
      "has been set to '" newlySetHandlerName "'."
    )
  }

  OnRestorePreviewHandler(fileExtension)
  {
    backupGuid := this.__RestoreBackup(fileExtension)
    
    if (this.previewHandlers.Has(backupGuid))
    {
      originalPreviewHandlerName := this.previewHandlers[backupGuid]
    }
    else
    {
      originalPreviewHandlerName := Format(PreviewHandlerManager
          .PREVIEW_HANDLER_TEXT.% backupGuid ? "UNKNOWN" : "NONE" %, 
          backupGuid)
    }

    this.gui["GUICurrentPreviewHandler"].Value := originalPreviewHandlerName

    ; Enable only when the restored handler differs from that selected in the dropdownlist
    this.gui["GUIBind"].Enabled := backupGuid 
        !== this.__DropdownPreviewHandlerGuid

    this.gui["GUIRestore"].Enabled := false

    msgbox (backupGuid ?
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
    this.gui["GUIRestore"].Enabled := this.gui["GUIBackup"].Value
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
  __CreateBackup(fileExtension, backupPayload)
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
      this.backupManager.Create(fileExtension, backupPayload, backupFormatNames*)
    }
  }

  ; Restore extension's original handler
  __RestoreBackup(fileExtension)
  {
    handlerBackup := this.backupManager.Retrieve(fileExtension, 
        this.chosenBackupFormats.Default)

    if (handlerBackup.guid)
    {
      ; Extension initially had a preview handler associated with it

      regwrite handlerBackup.guid, "REG_SZ", handlerBackup.registryBranch
    }
    else
    {
      ; Extension initially had no existing preview handler

      try
      {
        regdeletekey handlerBackup.registryBranch
      }
      catch Error as e
      {
        outputdebug Format("Unexpected app state: Registry key '{1}' could't be deleted: {2}", 
            handlerBackup.registryBranch, e.Message)
      }

      this.__RegRemoveShellExIfEmpty(handlerBackup.registryBranch)
    }

    return handlerBackup.guid
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

  __RegRemoveShellExIfEmpty(branch)
  {
    if (!instr(branch, "ShellEx"))
    {
      throw Error("Invalid registry branch without ShellEx key.", , branch)
    }

    ; Remove also "ShellEx" key if empty
    shellExKey := regexreplace(branch, "\\[^\\]+$")
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

  __RegIsWriteProtected(branch)
  {
    try
    {
      currentValue := regread(branch)
    }
    catch
    {
      ; No default value set
    }
    
    try
    {
      regwrite(currentValue ?? "", "REG_SZ", branch)

      if (!isset(currentValue))
      {
        ; Value was not set before → remove newly set empty string value

        regdeletekey(branch)
        try
        {
          this.__RegRemoveShellExIfEmpty(branch)
        }
        catch
        {
          ; Parent key can't be removed, whatever
        }
      }
    }
    catch Error as registryError
    {
      ; Can't be written to
      if (instr(registryError.Message, "Access is denied"))
      {
        return true
      }
      else
      {
        throw registryError ; Yet unidentified reason
      }
    }

    return false
  }

  __GetPreviewHandlerRegistryLocation(fileExtension)
  {
    static WINAPI :=
    {
      ERROR_SUCCESS: 0,
      HKEY_CLASSES_ROOT: 0x80000000,
      KEY_READ: 0x20019
    }, hKey := 0

    result :=
    {
      primaryLocationCandidate:
      {
        branch: "",
        isWriteProtected: false
      },
      secondaryLocationCandidate:
      {
        branch: "",
        isWriteProtected: false
      },
      actualLocationCandidate:
      {
        branch: "",
        isFound: true
      }
    }

    ; HKCR\.reg\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
    result.secondaryLocationCandidate.branch := Format(
        PreviewHandlerManager.REGISTRY_KEYNAME_FORMAT, fileExtension)

    ; HKCR\.reg
    delegatePrimaryKeyLocation := regexreplace(
        result.secondaryLocationCandidate.branch, "\\ShellEx.*$")

    ; regfile
    delegatePrimaryKeyName := regread(delegatePrimaryKeyLocation, , 0)

    if (delegatePrimaryKeyName)
    {
      ; HKCR\regfile\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
      result.primaryLocationCandidate.branch := Format(strreplace(
          PreviewHandlerManager.REGISTRY_KEYNAME_FORMAT, "."), delegatePrimaryKeyName)
      
      if (keyExists := WINAPI.ERROR_SUCCESS == DllCall("advapi32\RegOpenKeyEx", 
          "ptr", WINAPI.HKEY_CLASSES_ROOT, 
          "str", regexreplace(result.primaryLocationCandidate.branch, "^[A-Z_]+\\"),
          "int", 0, 
          "int", WINAPI.KEY_READ,
          "ptr*", hKey))
      {
        result.primaryLocationCandidate.isWriteProtected := this.__RegIsWriteProtected(
            result.primaryLocationCandidate.branch)

        if (!result.primaryLocationCandidate.isWriteProtected)
        {
          result.actualLocationCandidate.branch := result.primaryLocationCandidate.branch
          return result
        }
      }
    }

    ; HKCR\.reg\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
    if (keyExists := WINAPI.ERROR_SUCCESS == DllCall("advapi32\RegOpenKeyEx", 
        "ptr", WINAPI.HKEY_CLASSES_ROOT, 
        "str", regexreplace(result.secondaryLocationCandidate.branch, "^[A-Z_]+\\"),
        "int", 0, 
        "int", WINAPI.KEY_READ,
        "ptr*", hKey))
    {
      result.secondaryLocationCandidate.isWriteProtected := this.__RegIsWriteProtected(
          result.secondaryLocationCandidate.branch)

      if (!result.secondaryLocationCandidate.isWriteProtected)
      {
        result.actualLocationCandidate.branch := result.secondaryLocationCandidate.branch
        return result
      }
    }

    ; Neither location has a preview handler set OR either/both has/have it set but unwritable

    result.actualLocationCandidate.isFound := false

    if (result.primaryLocationCandidate.branch != "")
    {
      ; HKCR\regfile
      primaryLocationRoot := regexreplace(
          result.primaryLocationCandidate.branch, "\\ShellEx.*$")
    }

    if ((keyExists := isset(primaryLocationRoot) && 
        WINAPI.ERROR_SUCCESS == DllCall("advapi32\RegOpenKeyEx", 
            "ptr", WINAPI.HKEY_CLASSES_ROOT, 
            "str", regexreplace(primaryLocationRoot, "^[A-Z_]+\\"),
            "int", 0, 
            "int", WINAPI.KEY_READ,
            "ptr*", hKey)) &&
        !result.primaryLocationCandidate.isWriteProtected)
    {
      ; Primary location root key exists and likely won't be write protected
      result.actualLocationCandidate.branch := result.primaryLocationCandidate.branch
    }
    else
    {
      ; Secondary location may still be write protected, but we're out of options
      result.actualLocationCandidate.branch := result.secondaryLocationCandidate.branch
    }

    return result
  }

  __MenuCheckUpdate(url)
  {
    try
    {
      script.Update(url)
    }
    catch Any as e
    {
      if (e.HasOwnProp("msg"))
      {
        outputdebug Format("Update check ended with the code {1}, message: {2}", e.code, e.msg)
        
        if (e.code == 6) ; ERR_CURRENTVERSION
        {
          msgbox e.msg
        }
      }
      else
      {
        throw
      }
    }
  }

  __Throw(err)
  {
    throw err
  }

  static __GlobalExceptionHandler(*) ;(error, mode)
  {
    msgbox "An unhandled exception occurred."

    return 1
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
    get
    {
      result :=
      {
        guid: 0
      }

      registryKeyInfo := this.__GetPreviewHandlerRegistryLocation(fileExtension)
      
      if (registryKeyInfo.primaryLocationCandidate.isWriteProtected &&
          ((!registryKeyInfo.actualLocationCandidate.isFound) ||
          regread(registryKeyInfo.actualLocationCandidate.branch, , "") == ""))
      {
        ; At least .htm and .html extensions in Windows 10 will choose this logical path.
        ; Primary location is write protected, and Windows Explorer will actually use it
        ; as the secondary fallback location. The preview handler specified in 
        ; HKCR\.[ext]\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f} will be prioritized 
        ; to display file preview, unless the branch doesn't exist or the default value 
        ; is empty.

        sourceRegistryBranch := registryKeyInfo.primaryLocationCandidate.branch
      }
      else if (registryKeyInfo.actualLocationCandidate.isFound)
      {
        ; Other cases than those mentioned above, as long as a pre-existing registry 
        ; location for a preview handler was found

        sourceRegistryBranch := registryKeyInfo.actualLocationCandidate.branch
        
      }

      if (isset(sourceRegistryBranch))
      {
        extensionGuid := regread(sourceRegistryBranch, , "")

        if (extensionGuid != "")
        {
          result.guid := extensionGuid
        }
      }
      
      ; A preview handler may be set up in the primary location
      ; HKCR\{.[fileExtension] (Default) value}\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
      ; or in the secondary location
      ; HKCR\.[fileExtension]\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f}
      ; while, for whatever reason, not actually matching any available handler defined in
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
}