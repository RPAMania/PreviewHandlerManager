#include "IBackup.ahk"

class RuntimeMemoryBackup extends IBackup
{
  backup := Map()

  ; ============================================================
  ; Public methods
  ; ============================================================

    Create(uniqueBackupId, valueToBackup) => this.backup[uniqueBackupId] := valueToBackup

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