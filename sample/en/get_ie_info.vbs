' Building a reflector
Set reflector = CreateObject("ComUtilityLib.Reflector")

' Building IE object
Set ie = CreateObject("InternetExplorer.Application")

' Get a list of interfaces of IE object
WScript.Echo "# List of IE component interfaces"
For Each item In reflector.ScanIf(ie)
    WScript.Echo(" * " & item.Name)
Next
WScript.Echo

' List of IE object properties
WScript.Echo("# List of IE component properties")
For Each item In reflector.ScanDispIf(ie).Properties
    WScript.Echo(" * " + item.Name)
Next
WScript.Echo

' Search for ProgID that can construct IE objects
WScript.Echo("# List of IE component CLSID and ProgID")
For Each item In reflector.SearchCoClass(ie)
    WScript.Echo(" * " & item.CLSID & ", " & item.ProgID)
Next
WScript.Echo

