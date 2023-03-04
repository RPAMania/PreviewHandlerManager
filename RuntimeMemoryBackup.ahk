#include "IBackup.ahk"

class RuntimeMemoryBackup extends IBackup
{
  backup := Map()

  ; ============================================================
  ; Public methods
  ; ============================================================

    _Create(uniqueBackupId, valueToBackup) => this.backup[uniqueBackupId] := valueToBackup

    _Retrieve(uniqueBackupId)
    {
      if (!this._IsAlreadyCreated(uniqueBackupId))
      {
        this.__ThrowNonExistent(uniqueBackupId)
      }

      return this.backup[uniqueBackupId]
    }

    _Delete(uniqueBackupId)
    {
      if (!this._IsAlreadyCreated(uniqueBackupId))
      {
        this.__ThrowNonExistent(uniqueBackupId)
      }

      this.backup.Delete(uniqueBackupId)
    }

    _IsAlreadyCreated(uniqueBackupId) => this.backup.Has(uniqueBackupId)
}