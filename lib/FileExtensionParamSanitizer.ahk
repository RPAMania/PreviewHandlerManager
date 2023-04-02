class FileExtensionParamSanitizer
{
  static DotlessExtension[filePath] => regexreplace(filePath, "^\.")

  ; Inject file extension param auto-sanitization into parametrized GUI method calls
  __New()
  {
    proto := this

    ; Only convert until the least derived own custom type
    while ((proto := proto.base) !== FileExtensionParamSanitizer.Prototype)
    {
      for name in proto.OwnProps()
      {
        if (name ~= "^__" || ; Ignore double-underscore (private) methods
            !proto.HasMethod(name) || ; Ignore non-method properties
            proto.%name%.MinParams == 1) ; Ignore methods with no explicit param(s)
        {
          continue
        }

        ; MUST do injection in a separate function call instead of directly inside the 
        ; loop body, because only by forcing limited scope on the variable holding the 
        ; original method's func object will each fat-arrow capture the correct original 
        ; method, instead of the one retrieved by the very last iteration.
        ; "Each call to the outer function [i.e. __InjectParamSanitization instead of __New] 
        ; creates new closures, distinct from any previous calls" 
        ; @ https://www.autohotkey.com/docs/v2/Functions.htm#closures
        this.__InjectParamSanitization(proto, name)
      }
    }
  }

  ; Tap into the original method call and sanitize the param before relaying the actual call
  __InjectParamSanitization(proto, name)
  {
    methodDescriptor := proto.GetOwnPropDesc(name)

    originalMethod := methodDescriptor.Call
    methodDescriptor.Call := (instance, explicitParams*) => (
        explicitParams[1] := FileExtensionParamSanitizer.DotlessExtension[explicitParams[1]],
        originalMethod(instance, explicitParams*))
    proto.DefineProp(name, methodDescriptor)
  }
}