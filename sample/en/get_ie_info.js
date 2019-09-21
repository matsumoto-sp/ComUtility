// Building a reflector
var reflector = WScript.CreateObject("ComUtilityLib.Reflector");

// Building IE object
var ie = WScript.CreateObject("InternetExplorer.Application");

// Get a list of interfaces of IE object
WScript.Echo("# List of IE component interfaces");
var ifs = reflector.ScanIf(ie);
var eIfs = new Enumerator(ifs);
while (!eIfs.atEnd()) {
    WScript.Echo(' * ' + eIfs.item().Name);
    eIfs.moveNext();
}
WScript.Echo('');

// List of IE object properties
WScript.Echo("# List of IE component properties");
var props = reflector.ScanDispIf(ie).Properties;
var eProps = new Enumerator(props);
while (!eProps.atEnd()) {
    WScript.Echo(' * ' + eProps.item().Name);
    eProps.moveNext();
}
WScript.Echo('');

// Search for ProgID that can construct IE objects
WScript.Echo("# List of IE component CLSID and ProgID");
var coclasses = reflector.SearchCoClass(ie);
var eCoClasses = new Enumerator(coclasses);
while (!eCoClasses.atEnd()) {
    WScript.Echo(' * ' + eCoClasses.item().CLSID + ', ' + eCoClasses.item().ProgID);
    eCoClasses.moveNext();
}
WScript.Echo('');
