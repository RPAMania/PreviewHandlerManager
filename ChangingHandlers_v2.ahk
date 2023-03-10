/*
  NOTES BY T ON MAR 2, 23
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
      * Won't take into account a PersistentHandler that, if set, will override 
        extension's ShellEx handler
        > E.g. HKCR\.reg\PersistentHandler == {5e941d80-bf96-11cd-b579-08002b30bfeb}, 
          which refers to HKCR\CLSID\{5e941d80-bf96-11cd-b579-08002b30bfeb} with the 
          default value "Plain Text persistent handler"
    - "Use backups" checkbox
      * New backups won't not created while unticked
      * If unticked after a backup has been created, restoring the backup will be 
        disabled until ticked back on
      * Ticking/unticking while backups have already been created during a 
        session may possibly have unexpected consequences (not thoroughly tested)

  TODO
    - Allow registering new preview handlers?
    - Display the name of the original (backed-up) preview handler of an extension?
    - Backup feature de/activation should probably be more persistent than a freely 
      interactable checkbox in the main GUI, or at least more clearly defined in 
      terms of what happens if it's enabled/disabled on the fly.
*/

#requires autohotkey 2.0-beta+
#singleinstance ignore
#include "lib\PreviewHandlerManager.ahk"

if (!A_IsAdmin) { ;http://ahkscript.org/docs/Variables.htm#IsAdmin
  Run "*RunAs `"" A_ScriptFullPath "`""  ; Requires v1.0.92.01+
  ExitApp
}

if (!a_iscompiled)
{
  hotkey "F10", (*) => reload()
}

PreviewHandlerManager().Show()