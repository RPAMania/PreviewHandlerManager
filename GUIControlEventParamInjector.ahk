#include "FileExtensionParamSanitizer.ahk"

class GUIControlEventParamInjector extends FileExtensionParamSanitizer
{
  ; Prepare the extension param and redirect the call to
  ; the derived single-underscore method implementation
  __Call(methodName, params)
  {
    if (this.HasMethod("_" methodName) && substr(methodName, 1, 2) == "On")
    {
      fileExtension := FileExtensionParamSanitizer.DotlessExtensionFromPath[params[1]]

      return this.%"_" methodName%(fileExtension)
    }
    
    this.__ThrowMethodError(methodName)
  }
}