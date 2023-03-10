#include "RuntimeMemoryBackup.ahk"
#include "RegistryFileBackup.ahk"

class BackupManager
{
  class BackupFormat
  {
    static  RuntimeMemory := "RuntimeMemory", 
            RegistryFile := "RegistryFile"
            ; SqLiteDb: "SQLiteDb"
    
    ; Allow iterating with for ... in
    static __Enum(variableCount)
    {
      return this()
    }

    Call(varRefParams*)
    {
      if (varRefParams.Length < 1)
      {
        return false
      }

      getNext := false

      for backupFormatType, backupFormatName in BackupManager.BackupFormat.OwnProps()
      {
        ; Skip class prototype object
        if (backupFormatName == BackupManager.BackupFormat.Prototype)
        {
          continue
        }

        if (!this.HasProp("currentEnumBackupType"))
        {
          this.currentEnumBackupType := backupFormatType

          if (varRefParams.Has(1))
          {
            if (!varRefParams.Has(2))
            {
              ; Only first enum value requested
              %varRefParams[1]% := backupFormatName
            }
            else
            {
              %varRefParams[1]% := backupFormatType
            }
          }

          if (varRefParams.Has(2))
          {
            %varRefParams[2]% := backupFormatName
          }

          return true
        }
        else if (!getNext)
        {
          if (this.currentEnumBackupType == backupFormatType)
          {
            getNext := true
          }
        }
        else
        {
          this.currentEnumBackupType := backupFormatType
          
          if (varRefParams.Has(1))
          {
            if (!varRefParams.Has(2))
            {
              ; Only first enum value requested
              %varRefParams[1]% := backupFormatName
            }
            else
            {
              %varRefParams[1]% := backupFormatType
            }
          }

          if (varRefParams.Has(2))
          {
            %varRefParams[2]% := backupFormatName
          }

          return true
        }
      }

      return false
    }
  }

  backups := Map()

  __New(requestedBackupFormats)
  {
    for backupFormat, backupOption in requestedBackupFormats
    {
      switch (backupFormat)
      {
        case BackupManager.BackupFormat.RuntimeMemory:
        {
          this.backups[BackupManager.BackupFormat.RuntimeMemory] := RuntimeMemoryBackup()
        }
        case BackupManager.BackupFormat.RegistryFile:
        {
          this.backups[BackupManager.BackupFormat.RegistryFile] := RegistryFileBackup(
              backupOption.keyNameFormat, backupOption.fileNameFormat, 
              backupOption.hourTimeFormat, backupOption.guidToNameCallback)
        }
      }
    }
  }
  
  ; ============================================================
  ; Public methods
  ; ============================================================

    Create(fileExtension, valueToBackup, backupFormats*)
    {
      for , backupFormat in backupFormats
      {
        this.__ValidateBackupFormat(backupFormat)
        
        if (!this.__IsInUse(backupFormat))
        {
          this.__ThrowNotInUse(backupFormat)
        }

        this.backups[backupFormat].Create(fileExtension, valueToBackup)
      }
    }

    Retrieve(fileExtension, backupFormat)
    {
      this.__ValidateBackupFormat(backupFormat)

      if (!this.__IsInUse(backupFormat))
      {
        this.__ThrowNotInUse(backupFormat)
      }

      return this.backups[backupFormat].Retrieve(fileExtension)
    }

    IsAlreadyCreatedForSession(fileExtension, backupFormat)
    {
      this.__ValidateBackupFormat(backupFormat)
      
      return this.__IsInUse(backupFormat)
          && this.backups[backupFormat].IsAlreadyCreated(fileExtension)
    }

  ; ============================================================
  ; Private methods
  ; ============================================================
  
    __IsInUse(backupFormat)
    {
      return this.backups.Has(backupFormat)
    }

    __ValidateBackupFormat(backupFormat)
    {
      for supportedBackupFormatName in BackupManager.BackupFormat
      {
        if (backupFormat == supportedBackupFormatName)
        {
          return
        }
      }

      throw ValueError("Unknown BackupFormat value.", -2, backupFormat)
    }

    __ThrowNotInUse(backupFormat)
    {
      throw UnsetItemError(Format("Backup format {1} not in use.", backupFormat), -2)
    }
}