// リフレクターの構築
var reflector = WScript.CreateObject("ComUtilityLib.Reflector");

// IEオブジェクトの構築
var ie = WScript.CreateObject("InternetExplorer.Application");

// IE オブジェクトが持つインターフェースの一覧を得る
WScript.Echo("# IEコンポーネント・インターフェース一覧");
var ifs = reflector.ScanIf(ie);
var eIfs = new Enumerator(ifs);
while (!eIfs.atEnd()) {
    WScript.Echo(' * ' + eIfs.item().Name);
    eIfs.moveNext();
}
WScript.Echo('');

// IE オブジェクトのプロパティ一覧
WScript.Echo("# IEコンポーネント・プロパティ一覧");
var props = reflector.ScanDispIf(ie).Properties;
var eProps = new Enumerator(props);
while (!eProps.atEnd()) {
    WScript.Echo(' * ' + eProps.item().Name);
    eProps.moveNext();
}
WScript.Echo('');

// IE オブジェクトを構築できる ProgID を検索
WScript.Echo("# IEコンポーネント・CLSID, ProgID 一覧");
var coclasses = reflector.SearchCoClass(ie);
var eCoClasses = new Enumerator(coclasses);
while (!eCoClasses.atEnd()) {
    WScript.Echo(' * ' + eCoClasses.item().CLSID + ', ' + eCoClasses.item().ProgID);
    eCoClasses.moveNext();
}
WScript.Echo('');
