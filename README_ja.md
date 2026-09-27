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
    Set coll = RetrieveAllControlsInZOrderDfs(UserForm1)
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
    Set coll = RetrieveAllControlsInZOrderDfs(UserForm1)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    ws.Cells.Clear
    ws.Cells(1, 1) = "Type"
    ws.Cells(1, 2) = "Name"
    ws.Cells(1, 3) = "Z-Order Index"
    ws.Cells(1, 4) = "Z-Order Hierarchy"
    ws.Cells(1, 5) = "Depth"
    ws.Cells(1, 6) = "Parent Type"
    ws.Cells(1, 7) = "Parent Name"
    For Each item In coll
        ws.Cells(i, 1) = TypeName(item(0))
        ws.Cells(i, 2) = item(1)
        ws.Cells(i, 3) = item(2)
        ws.Cells(i, 4) = "'" & item(3)
        ws.Cells(i, 5) = item(4)
        ws.Cells(i, 6) = TypeName(item(0).Parent)
        ws.Cells(i, 7) = item(0).Parent.Name
        i = i + 1
    Next item
    ws.Cells.EntireColumn.AutoFit
End Sub
```
- 個別のコントロールのZオーダーを取得
```
Sub Test3()
    Debug.Print RetrieveCtrlZOrderIndex(UserForm1.CommandButton1)
    Debug.Print RetrieveCtrlZOrderHierarchy(UserForm1.CommandButton1)
End Sub
```
   > Note: 
   > - `UserForm1`および`UserForm1.CommandButton1`の部分は実際の対象のユーザーフォームまたはコントロールのオブジェクト名に変えて実行してください。
   > - `RetrieveCtrlZOrderIndex`と`RetrieveCtrlZOrderHierarchy`は内部的に`RetrieveAllControlsInZOrderDfs`を使用しており呼び出すたびに一旦すべてのコントロールを取得しているためループ内で使用するとパフォーマンスが著しく低下します、ループで使用する必要がある場合は`RetrieveAllControlsInZOrderDfs`で取得した`Collection`内でループ処理を実行してください。

## 関数一覧
[RetrieveAllControlsInZOrderDfs](#RetrieveAllControlsInZOrderDfs)  
[RetrieveCtrlZOrderIndex](#RetrieveCtrlZOrderIndex)  
[RetrieveCtrlZOrderHierarchy](#RetrieveCtrlZOrderHierarchy)  

### `RetrieveAllControlsInZOrderDfs`

深さ優先探索（DFS）を用いて、UserForm内のすべてのコントロールをZオーダー順に巡回・取得します。

#### シグネチャ
```vba
Public Function RetrieveAllControlsInZOrderDfs(ByVal objUserForm As MsForms.UserForm) As Collection
```

#### 引数
| 引数名 | 型 | 説明 |
| :--- | :--- | :--- |
| `objUserForm` | `MSForms.UserForm` | スキャン対象となるトップレベルの UserForm オブジェクト。 |

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

#### エラー
| エラーコード | 説明 |
| :---: | :--- |
| `515` | 渡されたオブジェクトがトップレベル（ルート）の UserForm ではない場合に発生します。 |

---

### `RetrieveCtrlZOrderIndex`

指定されたコントロールの、直近の親コンテナ内におけるZオーダーインデックス（1始まり）を取得します。

#### シグネチャ
```vba
Public Function RetrieveCtrlZOrderIndex(ByVal ctrl As Object) As Long
```

#### 引数
| 引数名 | 型 | 説明 |
| :--- | :--- | :--- |
| `ctrl` | `Object` | 照会対象となるコントロールまたは UserForm オブジェクト。 |

#### 戻り値
- **型**: `Long`
- **説明**: 親コンテナ内における 1始まり のZオーダーインデックス。
  - 指定されたオブジェクト自体がトップレベルの UserForm である場合は `0` を返します。
  - コントロールが見つからない場合は `-1` を返します。

---

### `RetrieveCtrlZOrderHierarchy`

UserForm 内における対象コントロールの階層的なZオーダーパス文字列を取得します。

#### シグネチャ
```vba
Public Function RetrieveCtrlZOrderHierarchy(ByVal ctrl As Object) As String
```

#### 引数
| 引数名 | 型 | 説明 |
| :--- | :--- | :--- |
| `ctrl` | `Object` | 照会対象となるコントロールまたは UserForm オブジェクト。 |

#### 戻り値
- **型**: `String`
- **説明**: トップレベルのコンテナから対象コントロールまでの完全なZオーダーパスを表すハイフン区切りの文字列（例: `"1-3-2"`）。
  - 指定されたオブジェクト自体がトップレベルの UserForm である場合は `"0"` を返します。
  - コントロールが見つからない場合は 空文字 (`""`) を返します。

## 謝辞
本プロジェクトは [https://github.com/tarboh/VBA-UserForm-Z-Order-Tool](https://github.com/tarboh/VBA-UserForm-Z-Order-Tool) のフォークです。元の開発者およびコントリビューターに感謝いたします。
