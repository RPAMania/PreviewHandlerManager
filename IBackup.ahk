class IBackup extends FileExtensionParamSanitizer
{
  _Create(*)            => this.__ThrowNotImplemented(a_thisfunc)
  _Retrieve(*)          => this.__ThrowNotImplemented(a_thisfunc)
  _Delete(*)            => this.__ThrowNotImplemented(a_thisfunc)
  _IsAlreadyCreated(*)  => this.__ThrowNotImplemented(a_thisfunc)
  
  __ThrowNonExistent(uniqueBackupId) => this.__Throw(ValueError(
      Format("{1} for '{2}' does not exist.", type(this), uniqueBackupId), -3))

  __ThrowNotImplemented(methodName) => this.__ThrowMethodError(methodName)
  
  __Throw(err)
  {
    throw err
  }
}