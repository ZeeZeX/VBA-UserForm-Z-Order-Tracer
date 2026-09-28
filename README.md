# VBA-UserForm-Z-Order-Tracer
:jp:[日本語の説明はこちら](https://github.com/ZeeZeX/VBA-UserForm-Z-Order-Tracer/blob/main/README_ja.md)<br><br>
A VBA library for getting the Z-order (stacking order) of UserForm controls.

## Usage
Add `UF_Z_Order_Tracer.bas` to your VBA project and retrieve the Z-order of controls as shown below:
- Retrieve the Z-order of all controls
```
Sub Test1()
    Dim coll As Collection
    Dim item As Variant
    Set coll = GetAllCtrlsInZOrder(UserForm1)
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
    Set coll = GetAllCtrlsInZOrder(UserForm1)
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
- Retrieve the Z-order of individual controls
```
Sub Test3()
    Debug.Print GetCtrlZOrderIndex(UserForm1.CommandButton1)
    Debug.Print GetCtrlZOrderHierarchy(UserForm1.CommandButton1)
End Sub
```

>Note:
>
>-   Replace `UserForm1` and `UserForm1.CommandButton1` with your actual UserForm or control object names before running the code.
>
>-   `GetCtrlZOrderIndex` and `GetCtrlZOrderHierarchy` internally use `GetAllCtrlsInZOrder`, fetching all controls every time they are called. Using them inside a loop will significantly decrease performance. If you need to process controls within a loop, please iterate over the `Collection` retrieved by `GetAllCtrlsInZOrder`.

## Functions
[GetAllCtrlsInZOrder](#GetAllCtrlsInZOrder)  
[GetCtrlZOrderIndex](#GetCtrlZOrderIndex)  
[GetCtrlZOrderHierarchy](#GetCtrlZOrderHierarchy)  

### `GetAllCtrlsInZOrder`

Traverses and retrieves all controls in a UserForm in Z-Order using Depth-First Search (DFS).

#### Signature
```vba
Public Function GetAllCtrlsInZOrder(ByVal objUserForm As MsForms.UserForm) As Collection
```

#### Parameters
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `objUserForm` | `MSForms.UserForm` | The target UserForm object to scan. |

#### Return Value
- **Type**: `Collection`
- **Description**: A `Collection` containing 0-based `Variant` arrays for each control found. Each array element consists of the following structure:

| Index | Type | Description |
| :---: | :--- | :--- |
| `0` | `Object` | The Control object itself. |
| `1` | `String` | Control name. |
| `2` | `Long` | Z-Order index within its immediate parent container (1-based). |
| `3` | `String` | Hierarchical Z-Order path string (e.g., `"1-2-1"`). |
| `4` | `Long` | Nesting depth level starting from `1`. |

#### Errors Raised
| Error Code | Description |
| :---: | :--- |
| `515` | Thrown if the provided object is not the top-level (root) UserForm. |

---

### `GetCtrlZOrderIndex`

Retrieves the 1-based Z-Order index of a specific control within its immediate parent container.

#### Signature
```vba
Public Function GetCtrlZOrderIndex(ByVal ctrl As Object) As Long
```

#### Parameters
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `ctrl` | `Object` | The target control or UserForm object to query. |

#### Return Value
- **Type**: `Long`
- **Description**: The 1-based Z-Order index of the control within its parent container.
  - Returns `0` if the provided object is the top-level UserForm itself.
  - Returns `-1` if the control is not found.

---

### `GetCtrlZOrderHierarchy`

Retrieves the hierarchical Z-Order path string for a target control within a UserForm.

#### Signature
```vba
Public Function GetCtrlZOrderHierarchy(ByVal ctrl As Object) As String
```

#### Parameters
| Parameter | Type | Description |
| :--- | :--- | :--- |
| `ctrl` | `Object` | The target control or UserForm object to query. |

#### Return Value
- **Type**: `String`
- **Description**: A hyphen-delimited string representing the full Z-Order path from the top-level container down to the target control (e.g., `"1-3-2"`).
  - Returns `"0"` if the provided object is the top-level UserForm itself.
  - Returns an empty string (`""`) if the control is not found.

## Acknowledgements
This project is a fork of [https://github.com/tarboh/VBA-UserForm-Z-Order-Tool](https://github.com/tarboh/VBA-UserForm-Z-Order-Tool). Special thanks to the original authors and contributors.
