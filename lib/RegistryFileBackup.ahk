#include "IBackup.ahk"

class RegistryFileBackup extends IBackup
{
  static HourTimeFormat := { 12: "tt'_'hhmmss", 24: "HHmmss"}

  backup := Map()

  __New(fileNameFormat, hourTimeFormat, guidToNameCallback)
  {
    super.__New()
    
    this.__ValidateHourTimeFormat(hourTimeFormat)
    this.__SetBackupFileName(fileNameFormat, hourTimeFormat)

    this.GuidToNameCallback := guidToNameCallback
  }

  ; ============================================================
  ; Public methods
  ; ============================================================

    Create(uniqueBackupId, backupPayload)
    {
      ; Place every backup in the same .reg file

      if (!fileexist(this.backupFileName))
      {
        fileappend
        (
          "Windows Registry Editor Version 5.00
          
          ; Original registry values below
          
          "
        ), this.backupFileName, "`n"
      }

      existingPreviewHandlerGuid := regread(backupPayload.registryBranch, , 0)

      fileappend (existingPreviewHandlerGuid !== 0 ?
      (
        "; " this.GuidToNameCallback.Call(existingPreviewHandlerGuid) "
        [" backupPayload.registryBranch "]
        @=`"" existingPreviewHandlerGuid "`"
        
        "
      ) :
      (
        "[-" backupPayload.registryBranch "]
        
        "
      )), this.backupFileName, "`n"

      this.backup[uniqueBackupId] := a_scriptdir "\" this.backupFileName
    }

    Retrieve(uniqueBackupId)
    {
      if (!this.IsAlreadyCreated(uniqueBackupId))
      {
        this.__ThrowNonExistent(uniqueBackupId)
      }

      return this.backup[uniqueBackupId]
    }

    Delete(uniqueBackupId)
    {
      if (!this.IsAlreadyCreated(uniqueBackupId))
      {
        this.__ThrowNonExistent(uniqueBackupId)
      }

      registryFileFullPath := this.backup[uniqueBackupId]

      if (fileexist(registryFileFullPath))
      {
        try
        {
          filedelete registryFileFullPath
        }
        catch Error as e
        {
          msgbox Format("Error deleting backup file {1}: {2}", registryFileFullPath, e.Message)
        }
      }

      this.backup.Delete(uniqueBackupId)
    }

    IsAlreadyCreated(uniqueBackupId) => this.backup.Has(uniqueBackupId)

  ; ============================================================
  ; Private methods
  ; ============================================================

    ; Determine timestamp format used in the file name
    __SetBackupFileName(filenameFormat, hourTimeFormat)
    {
      static LOCALE_NAME_USER_DEFAULT := 0

      switch (hourTimeFormat)
      {
        case RegistryFileBackup.HourTimeFormat.12:
          ; When 12-hour timestamp format is requested, better use INVARIANT instead 
          ; of current user locale, because if 24-hour format is in effect, "tt" 
          ; will (likely) never translate into "AM"/"PM" but be left blank instead.
          LOCALE_NAME_INVARIANT := ""
        case RegistryFileBackup.HourTimeFormat.24:
        default:
          throw ValueError("Unsupported time format.", -3, hourTimeFormat)
      }

      charCount := 0
      currentTime := ""
      loop 2
      {
        if (a_index == 2)
        {
          varsetstrcapacity(&currentTime, charCount)
        }        

        shouldUse12HourFormat := hourTimeFormat == RegistryFileBackup.HourTimeFormat.12 
            && (charCount == 0 || 
                charCount == strlen(strreplace(RegistryFileBackup.HourTimeFormat.12, "'")) + 1)

        ; 1st call: Get required min length for storing the time format string
        ; 2nd call: Get time format string
        charCount := dllcall("GetTimeFormatEx", 
            isset(LOCALE_NAME_INVARIANT) ? "str" : "ptr", 
                LOCALE_NAME_INVARIANT ?? LOCALE_NAME_USER_DEFAULT,
            "int", 0,
            "ptr", 0,
            "str", RegistryFileBackup.HourTimeFormat.%shouldUse12HourFormat ? 12 : 24%,
            "str", currentTime,
            "int", charCount)
      }

      currentDate := formattime(unset, "yyyyMMdd")

      this.backupFileName := Format(filenameFormat, currentDate "_" currentTime)
    }

    __ValidateHourTimeFormat(hourTimeFormat)
    {
      for , supportedHourTimeFormat in RegistryFileBackup.HourTimeFormat.OwnProps()
      {
        if (hourTimeFormat == supportedHourTimeFormat)
        {
          return
        }
      }

      throw ValueError("Unknown HourTimeFormat value.", -3, hourTimeFormat)
    }
}