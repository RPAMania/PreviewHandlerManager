/*
  Notes by T on Mar 2, 2023:
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
  
  Apr 1, 2023:
    - Implemented handling of the alternate (=primary) preview handler location in the registry,
      HKCR\{HKCR\.ext\(Default) value}\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f},
      such as HKCR\regfile\ShellEx\{8895b1c6-b41f-4c1c-a562-0d564250836f} for .reg files
    - Improve validation and handling of file extension input by a user + apply custom red 
      background color to the file extension Edit field to indicate invalid input
  
  TODO:
    - Allow registering new preview handlers?
    - Display the name of the original (backed-up) preview handler of an extension? 
*/

#requires autohotkey v2.0
#singleinstance ignore
#include lib\PreviewHandlerManager.ahk

if (!A_IsAdmin) { ;http://ahkscript.org/docs/Variables.htm#IsAdmin
  Run "*RunAs `"" A_ScriptFullPath "`""  ; Requires v1.0.92.01+
  ExitApp
}

if (a_iscompiled)
{
  OnError PreviewHandlerManager.__GlobalExceptionHandler.Bind(PreviewHandlerManager)
}

PreviewHandlerManager().Show()