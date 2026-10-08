# VBA-UserForm-Z-Order-Tracer
🌎[English](https://github.com/ZeeZeX/VBA-UserForm-Z-Order-Tracer/blob/main/README.md)<br><br>
このライブラリは、VBAのユーザーフォームにおけるコントロールのZオーダー（重なり順）を取得するためのものです。

## 使用方法
`UF_Z_Order_Tracer.bas`をVBAプロジェクトに追加し、以下のようにコントロールのZオーダーを取得してください。
- 全コントロールのZオーダーを取得
```
Sub Test1()
    Dim coll As Collection
    Dim item As Variant
    Set coll = GetAllCtrlsInZOrder(UserForm1, zOrderType:="Internal")
    For Each item In coll
        Debug.Print TypeName(item(0)), item(1), item(2), item(3), item(4)
    Next item
End Sub
```
```
Sub Test2()
    Dim coll As Collection
    Dim item As Variant
    Dim i
    i = 2
    Set coll = GetAllCtrlsInZOrder(UserForm1, zOrderType:="Internal")
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    ws.Cells.Clear
    ws.Cells(1, 1) = "Type"
    ws.Cells(1, 2) = "Name"
    ws.Cells(1, 3) = "Z-Order Index"
    ws.Cells(1, 4) = "Z-Order Hierarchy"
    ws.Cells(1, 5) = "Depth"
    ws.Cells(1, 6) = "Address"
    ws.Cells(1, 7) = "Caption"
    ws.Cells(1, 8) = "Parent Type"
    ws.Cells(1, 9) = "Parent Name"
    ws.Cells(1, 10) = "Parent Address"
    ws.Cells(1, 11) = "Parent Caption"
    Application.EnableEvents = False
    For Each item In coll
        ws.Cells(i, 1) = TypeName(item(0))
        ws.Cells(i, 2) = item(1)
        ws.Cells(i, 3) = item(2)
        ws.Cells(i, 4) = "'" & item(3)
        ws.Cells(i, 5) = item(4)
        ws.Cells(i, 6) = "0x" & Hex(ObjPtr(item(0)))

        ws.Cells(i, 8) = TypeName(item(0).Parent)
        ws.Cells(i, 9) = item(0).Parent.Name
        ws.Cells(i, 10) = "0x" & Hex(ObjPtr(item(0).Parent))
        
        On Error Resume Next
        ws.Cells(i, 7) = item(0).Caption
        ws.Cells(i, 11) = item(0).Parent.Caption
        On Error GoTo 0
        
        i = i + 1
    Next item
    Application.EnableEvents = True
    ws.Cells.EntireColumn.AutoFit
End Sub
```
- 個別のコントロールのZオーダーを取得
```
Sub Test3()
    Debug.Print GetCtrlZOrderIndex(UserForm1.CommandButton1)
    Debug.Print GetCtrlZOrderHierarchy(UserForm1.CommandButton1)
End Sub
```
```
Sub Test4()
    Dim coll As Collection
    Dim ctrl As Object
    Dim key As String
    Set coll = GetAllCtrlsInZOrder(UserForm1)
    Set ctrl = UserForm1.CommandButton1
    key = Hex(ObjPtr(ctrl))
    Debug.Print coll(key)(2)
    Debug.Print coll(key)(3)
End Sub
```

   > Note: 
   > - `UserForm1`および`UserForm1.CommandButton1`の部分は実際の対象のユーザーフォームまたはコントロールのオブジェクト名に変えて実行してください。
   > - `GetCtrlZOrderIndex` および `GetCtrlZOrderHierarchy` は内部的に `GetAllCtrlsInZOrder` を使用しており、呼び出されるたびにすべてのコントロールを取得します。そのため、ループ内で使用するとパフォーマンスが著しく低下します。ループ内でコントロールを処理する必要がある場合は、あらかじめ `GetAllCtrlsInZOrder` で `Collection` を一度だけ取得し、`Hex(ObjPtr(ctrl))` をキーとして参照（例：`coll(Hex(ObjPtr(UserForm1.CommandButton1)))`）してください。

## 関数一覧
[GetAllCtrlsInZOrder](#GetAllCtrlsInZOrder)  
[GetCtrlZOrderIndex](#GetCtrlZOrderIndex)  
[GetCtrlZOrderHierarchy](#GetCtrlZOrderHierarchy)  

### `GetAllCtrlsInZOrder`

深さ優先探索（DFS）を用いて、UserForm内のすべてのコントロールをZオーダー順に巡回・取得します。

#### シグネチャ
```vba
Public Function GetAllCtrlsInZOrder(ByVal objUserForm As MsForms.UserForm, Optional ByVal zOrderType As String = "Internal") As Collection
```

#### 引数
| 引数名 | 型 | 説明 |
| :--- | :--- | :--- |
| `objUserForm` | `MSForms.UserForm` | スキャン対象となるトップレベルの UserForm オブジェクト。 |
| `zOrderType` | `String` | （省略可能） MultiPage 内の Page コントロールに対する Z オーダーの評価方法を指定します。デフォルトは `"Internal"` です。<br>•`"Internal"`: MultiPage 内の Page コントロールに対して内部的な Z オーダーの取得を優先します。見た目のページ順序が移動されている場合でも、追加された順番で取得されます。<br>•`"Visual"`: 見た目のタブ順序を優先し、左側のタブから順に Z オーダーを割り当てます。ページの並べ替えが反映されます。<br>•`"VisualKeepZ"`: `"Internal"` と同じ Z オーダー値を保持しつつ、返されるコレクションを `"Visual"` と同じ順序にソートします。 |
| `bfsSort` | `Boolean` | （省略可能） 返されるコレクションを幅優先順（BFS）にソートするかどうかを指定します。<br>`True`の場合コントロールをネストの深さに基づいて浅い順から深い順に並べ替え、同じ深さのコントロール間では元の順序を維持します。<br>デフォルトは `False`（深さ優先順）です。 |

#### 戻り値
- **型**: `Collection`
- **説明**: 検出された各コントロールの情報（`Variant` 配列）を格納した `Collection`。各配列の要素構造は以下の通りです：

| インデックス | 型 | 説明 |
| :---: | :--- | :--- |
| `0` | `Object` | コントロールオブジェクト本体 |
| `1` | `String` | コントロール名 |
| `2` | `Long` | 直近の親コンテナ内におけるZオーダーインデックス（1始まり） |
| `3` | `String` | 階層的なZオーダー文字列（例: `"1-2-1"`） |
| `4` | `Long` | ネストの深さレベル（ルート直下を `1` とする） |

>Note:
>返されるコレクション内のアイテムは、`Hex(ObjPtr(ctrl))` をキーとして取得できます。
>コントロールを保持する変数には、`MSForms.Control` や特定のコントロール型ではなく、`Object` または `Variant` として宣言してください。`MSForms.Control` に代入すると、`ObjPtr` が異なるポインタアドレスを返す場合があります。（このポインタの不一致は、特に MultiPage 内の Page コントロールで確認されています）


#### エラー
| エラーコード | 説明 |
| :---: | :--- |
| `515` | 渡されたオブジェクトがトップレベル（ルート）の UserForm ではない場合に発生します。 |
| `516` | 無効またはサポートされていない `zOrderType` 文字列が指定された場合に発生します。 |

---

### `GetCtrlZOrderIndex`

指定されたコントロールの、直近の親コンテナ内におけるZオーダーインデックス（1始まり）を取得します。

#### シグネチャ
```vba
Public Function GetCtrlZOrderIndex(ByVal ctrl As Object, Optional ByVal zOrderType As String = "Internal") As Long
```

#### 引数
| 引数名 | 型 | 説明 |
| :--- | :--- | :--- |
| `ctrl` | `Object` | 照会対象となるコントロールまたは UserForm オブジェクト。 |
| `zOrderType` | `String` | （省略可能）MultiPage 内の Page コントロールに対する Z オーダーの評価方法を指定します。デフォルトは `"Internal"` です。<br>•`"Internal"`: MultiPage 内の Page コントロールに対して内部的な Z オーダーの取得を優先します。見た目のページ順序が移動されている場合でも、追加された順番で取得されます。<br>•`"Visual"`: 見た目のタブ順序を優先し、左側のタブから順に Z オーダーを割り当てます。ページの並べ替えが反映されます。|

#### 戻り値
- **型**: `Long`
- **説明**: 親コンテナ内における 1始まり のZオーダーインデックス。
  - 指定されたオブジェクト自体がトップレベルの UserForm である場合は `0` を返します。
  - コントロールが見つからない場合は `-1` を返します。

#### エラー
| エラーコード | 説明 |
| :---: | :--- |
| `516` | 無効またはサポートされていない `zOrderType` 文字列が指定された場合に発生します。 |

---

### `GetCtrlZOrderHierarchy`

UserForm 内における対象コントロールの階層的なZオーダーパス文字列を取得します。

#### シグネチャ
```vba
Public Function GetCtrlZOrderHierarchy(ByVal ctrl As Object, Optional ByVal zOrderType As String = "Internal") As String
```

#### 引数
| 引数名 | 型 | 説明 |
| :--- | :--- | :--- |
| `ctrl` | `Object` | 照会対象となるコントロールまたは UserForm オブジェクト。 |
| `zOrderType` | `String` | （省略可能）MultiPage 内の Page コントロールに対する Z オーダーの評価方法を指定します。デフォルトは `"Internal"` です。<br>•`"Internal"`: MultiPage 内の Page コントロールに対して内部的な Z オーダーの取得を優先します。見た目のページ順序が移動されている場合でも、追加された順番で取得されます。<br>•`"Visual"`: 見た目のタブ順序を優先し、左側のタブから順に Z オーダーを割り当てます。ページの並べ替えが反映されます。|

#### 戻り値
- **型**: `String`
- **説明**: トップレベルのコンテナから対象コントロールまでの完全なZオーダーパスを表すハイフン区切りの文字列（例: `"1-3-2"`）。
  - 指定されたオブジェクト自体がトップレベルの UserForm である場合は `"0"` を返します。
  - コントロールが見つからない場合は 空文字 (`""`) を返します。

#### エラー
| エラーコード | 説明 |
| :---: | :--- |
| `516` | 無効またはサポートされていない `zOrderType` 文字列が指定された場合に発生します。 |

## クレジット

本プロジェクトはtarboh氏の[VBA-UserForm-Z-Order-Tool](https://github.com/tarboh/VBA-UserForm-Z-Order-Tool) をベースに大幅な拡張・変更を加えたものです。
