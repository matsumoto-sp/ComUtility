' リフレクターの構築
Set reflector = CreateObject("ComUtilityLib.Reflector")

' IEオブジェクトの構築
Set ie = CreateObject("InternetExplorer.Application")

' IE オブジェクトが持つインターフェースの一覧を得る
WScript.Echo "# IEコンポーネント・インターフェース一覧"
For Each item In reflector.ScanIf(ie)
    WScript.Echo(" * " & item.Name)
Next
WScript.Echo

' IE オブジェクトのプロパティ一覧
WScript.Echo("# IEコンポーネント・プロパティ一覧")
For Each item In reflector.ScanDispIf(ie).Properties
    WScript.Echo(" * " + item.Name)
Next
WScript.Echo

' IE オブジェクトを構築できる ProgID を検索
WScript.Echo("# IEコンポーネント・CLSID, ProgID 一覧")
For Each item In reflector.SearchCoClass(ie)
    WScript.Echo(" * " & item.CLSID & ", " & item.ProgID)
Next
WScript.Echo

