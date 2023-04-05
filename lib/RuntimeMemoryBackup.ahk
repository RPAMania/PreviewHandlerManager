#include "IBackup.ahk"

class RuntimeMemoryBackup extends IBackup
{
  backup := Map()

  __New()
  {
    super.__New()
  }
  
  ; ============================================================
  ; Public methods
  ; ============================================================

    Create(uniqueBackupId, backupPayload) => this.backup[uniqueBackupId] := backupPayload

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

      this.backup.Delete(uniqueBackupId)
    }

    IsAlreadyCreated(uniqueBackupId) => this.backup.Has(uniqueBackupId)
}