# FunctionalVBA
Enables functional programming (to a large degree) in VBA via a `FunctionObject` class module. Lambda functions still elude VBA programmers, so target functions need to be defined elsewhere.

The `FunctionObject` stores information required to execute a function via `Application.Run` or `CallByName`. It effectively wraps a function into an object that can be passed around your code and called at a later point, similar to other functional programming languages.

The `WrapFunction` method should be called shortly after instantiating the `FunctionObject` to provide details as to which target function to wrap:
- `funcName As String` The function name that is to be wrapped and called
- `Optional owner As Object` May be provided if this wrapper is intended to call a function on another object (`CallByName` method). Defaults to `Nothing`
- `Optional callType as VbCallType` May be provided if the function object is intended to call a property `Let`, `Get` or `Set` method (`CallByName` method). Defaults to `vbMethod`

Arguments can optionally be pre-bound via the `BindArgs` method. These arguments are passed to the target function in addition to arguments specified when invoking the function. This is very handy if there is data that needs to be passed, but the code that actually invokes the function doesn't have access, scope, or know-about. The code that invokes the function may also pass additional parameters at invoke-time.

The `CallFunction` method is used to invoke the wrapped function. It accepts arguments that are to be passed to the wrapped function. If there is data bound by `BindArgs`, these are placed at the front of the function's parameter list followed by the arguments passed in `CallFunction`. Arguments are passed in the order they are provided on the parameter lists for `BindArgs` followed by `CallFunction`. The return value of `CallFunction` is the return value of the target function, if applicable.

There can only be up to 30 arguments provided to the target function between the bound args and invoke-time args combined.

The target function must be publicly visible. The target function must accept type `Variant` for all of it's parameters or a `Type Mismatch` error will occur.

## Usage Example
```VBA
Public Function Target(name As Variant, coll As Variant, max As Variant) As Variant
  Debug.Print "In Target | Name - " & name
  coll.Add "New item"
  
  Target = max + 1
End Function


Public Function Invoker(fo As FunctionObject) As Integer
  Invoker = fo.CallFunction(3) ' fo already has two bound args so 3 passes to max
End Function


Public Sub SetupFunctionObject()
  Dim fo As New FunctionObject
  Dim entryColl As New Collection
  
  Call fo.WrapFunction("Target")
  Call fo.BindArgs("Hello!", entryColl) ' "Hello!" passes to name, entryColl to coll
  
  Dim result as Integer
  result = Invoker(fo)
  
  Debug.Print "Result - " & result
  Debug.Print "Collection contains - " & entryColl(1)
End Sub
```
## Example Output
```
In Target | Name - Hello!
Result - 4
Collection contains - New item
```
