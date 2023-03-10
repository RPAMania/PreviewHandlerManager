class IBackup
{
  Create(*)            => this.__ThrowNotImplemented(a_thisfunc)
  Retrieve(*)          => this.__ThrowNotImplemented(a_thisfunc)
  Delete(*)            => this.__ThrowNotImplemented(a_thisfunc)
  IsAlreadyCreated(*)  => this.__ThrowNotImplemented(a_thisfunc)
  
  __ThrowNonExistent(uniqueBackupId) => this.__Throw(ValueError(
      Format("{1} for '{2}' does not exist.", type(this), uniqueBackupId), -3))

  __ThrowNotImplemented(methodName) => this.__Throw(MethodError(
      "Not implemented", -2, this.__MethodWithoutClassName(methodName)))

  __MethodWithoutClassName(methodName) => regexreplace(methodName, ".*?([^.]+)$", "$1")

  __Throw(err)
  {
    throw err
  }
}