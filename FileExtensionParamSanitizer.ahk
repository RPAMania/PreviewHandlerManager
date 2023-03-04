class FileExtensionParamSanitizer
{
  static DotlessExtensionFromPath[filePath] => regexreplace(filePath, "^.*\.")

  ; Prepare the extension param and redirect the call to 
  ; the derived single-underscore method implementation
  __Call(methodName, params)
  {
    if (this.HasMethod("_" methodName, params.Length) && params.Length)
    {
      params[1] := FileExtensionParamSanitizer.DotlessExtensionFromPath[params[1]]

      return this.%"_" methodName%(params*)
    }

    this.__ThrowMethodError(methodName)
  }

  __ThrowMethodError(methodName)
  {
    throw MethodError("Not implemented", -3, this.__MethodWithoutClassName(methodName))
  }

  __MethodWithoutClassName(methodName) => regexreplace(methodName, "^.*\._?([^.]+)$", "$1")
}