Attribute VB_Name = "UF_Z_Order_Tracer"

' ========================================================================================
' VBA-UserForm-Z-Order-Tracer v1.0.4
' https://github.com/ZeeZeX/VBA-UserForm-Z-Order-Tracer
' Copyright (c) 2026 ZeeZeX
' License: MIT
'
' Based originally on tarboh's VBA-UserForm-Z-Order-Tool with extensive changes/additions
' https://github.com/tarboh/VBA-UserForm-Z-Order-Tool
' Copyright (c) 2026 tarboh
' License: MIT
'
' https://opensource.org/licenses/MIT
' ========================================================================================

Option Explicit

#If VBA7 = 0 Then
    ' Define a placeholder LongPtr type for VBA6 or earlier.
    ' This allows code that uses LongPtr to compile in older VBA versions,
    ' even though those versions do not natively support the LongPtr type.
    Private Enum LongPtr
        [_]
    End Enum
#End If


#If VBA7 Then
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cArgs As Long, _
    ByRef rgvt As Integer, _
    ByRef rgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long
#Else
Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As Long, _
    ByVal oVft As Long, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cArgs As Long, _
    ByRef rgvt As Integer, _
    ByRef rgpvarg As Long, _
    ByRef pvargResult As Variant) As Long
#End If



Public Function GetAllCtrlsInZOrder(ByVal objUserForm As MSForms.UserForm, Optional ByVal zOrderType As String = "Internal", Optional bfsSort As Boolean = False) As Collection
    ''
    ' Traverses and retrieves all controls in a UserForm in Z-Order using Depth-First Search (DFS).
    '
    ' @param objUserForm [In] The target UserForm object (MsForms.UserForm) to scan.
    ' @param zOrderType  [In] Optional mode to determine how Z-Order is evaluated for Page controls inside a MultiPage.
    '                         Default is "Internal".
    '                           - "Internal": Prioritizes internal Z-Order retrieval for Page controls inside a MultiPage.
    '                                         Controls are retrieved in addition order, even if the visual page order has been moved.
    '                           - "Visual":   Prioritizes visual tab order, assigning Z-Order sequentially starting from the left tab.
    '                                         Reflects any reordering of pages.
    '                           - "VisualKeepZ": Retrieves Z-Order values identical to "Internal", but sorts the resulting collection
    '                                            in the same order as "Visual".
    '
    ' @param bfsSort     [In] Optional flag to sort the returned collection in breadth-first order (BFS).
    '                             If True, controls are ordered by nesting depth, from shallowest to deepest,
    '                             while preserving their original order within the same depth.
    '                             Default is False (depth-first order).
    '
    ' @return Collection A Collection containing Variant arrays for each control found.
    '                     Each array element consists of:
    '                       (0): Control object (Object)
    '                       (1): Control name (String)
    '                       (2): Z-Order index within its immediate container (Long)
    '                       (3): Hierarchical Z-Order string (e.g., "1-2-1") (String)
    '                       (4): Nesting depth level starting from 1 (Long)
    '
    '                     Items in the returned collection can be retrieved by key using `Hex(ObjPtr(ctrl))`.
    '                     Note: Variables holding controls should be declared as `Object` or `Variant` instead of
    '                     `MSForms.Control` or specific control types, as assigning to `MSForms.Control` may
    '                     cause `ObjPtr` to return a different pointer address. (this pointer mismatch has been
    '                     specifically confirmed with Page controls within a MultiPage)
    '
    ' @raises 515 Thrown if the provided object is not the top-level (root) UserForm.
    ' @raises 516 Thrown if an invalid or unsupported zOrderType string is specified.
    ''
    Dim result As Collection
    Dim ctrlsInternalOrder As Collection
    Dim ctrlsVisialOrder As Collection
    Dim item As Variant
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim key As String
    Dim tempColl As New Collection
    Dim tempColl2 As Collection
    Dim maxDepth As Long
    Dim depth As Long
    Dim i As Long
    Const q = """"
    Dim errMsg As String
    Set result = New Collection
    ' Raise an error if the passed object is not the root UserForm.
    If Not objUserForm Is GetUserFormObjectFromCtrl(objUserForm) Then
        Err.Raise Number:=515, Description:="Invalid UserForm"
    End If
    
    Select Case UCase$(Trim$(zOrderType))
    
        Case "INTERNAL"
            TraceZOrderRecursive objUserForm, 1, "", result, useVisualOrder:=False
            
        Case "VISUAL"
            TraceZOrderRecursive objUserForm, 1, "", result, useVisualOrder:=True
            
        Case "VISUALKEEPZ"
            Set ctrlsInternalOrder = New Collection
            Set ctrlsVisialOrder = New Collection
            TraceZOrderRecursive objUserForm, 1, "", ctrlsInternalOrder, useVisualOrder:=False
            TraceZOrderRecursive objUserForm, 1, "", ctrlsVisialOrder, useVisualOrder:=True
            
            For Each item In ctrlsVisialOrder
                key = Hex(ObjPtr(item(0)))
                arr = item
                arr2 = ctrlsInternalOrder(key)
                arr(2) = arr2(2)
                arr(3) = arr2(3)
                result.Add arr, key
            Next item
            
        Case Else
            errMsg = "[zOrderType] Invalid value: " & q & zOrderType & q & vbLf & "Supported values are " & q & Join(Array("Internal", "Visual", "VisualkeepZ"), q & ", " & q) & q
            Err.Raise Number:=516, Description:=errMsg
    End Select
    
    If bfsSort Then
        maxDepth = 0
        For Each item In result
            depth = item(4)
            If depth > maxDepth Then
                maxDepth = depth
            End If
            key = "key" & CStr(depth)
            On Error Resume Next
            tempColl.Add New Collection, key
            On Error GoTo 0
            tempColl(key).Add item
        Next
        Set result = New Collection
        For i = 1 To maxDepth
            key = "key" & CStr(i)
            Set tempColl2 = Nothing
            On Error Resume Next
            Set tempColl2 = tempColl(key)
            On Error GoTo 0
            If Not tempColl2 Is Nothing Then
                For Each item In tempColl2
                    result.Add item, Hex(ObjPtr(item(0)))
                Next
            End If
        Next i
    End If
    
    Set GetAllCtrlsInZOrder = result
    
End Function

Public Function GetCtrlZOrderIndex(ByVal ctrl As Object, Optional ByVal zOrderType As String = "Internal") As Long
    ' Retrieves the 1-based Z-Order index of a specific control within its immediate parent container.
    '
    ' @param ctrl [In] The target control or UserForm object to locate.
    ' @param zOrderType  [In] Optional mode to determine how Z-Order is evaluated for Page controls inside a MultiPage.
    '                         Default is "Internal".
    '                           - "Internal": Prioritizes internal Z-Order retrieval for Page controls inside a MultiPage.
    '                                         Controls are retrieved in addition order, even if the visual page order has been moved.
    '                           - "Visual":   Prioritizes visual tab order, assigning Z-Order sequentially starting from the left tab.
    '                                         Reflects any reordering of pages.
    ' @return Long The 1-based Z-Order index of the control within its parent container.
    '              Returns 0 if the provided object is the top-level UserForm itself.
    '              Returns -1 if the control is not found.
    '
    ' @raises 516 Thrown if an invalid or unsupported zOrderType string is specified.
    Dim result As Long
    Dim key As String
    result = -1
    Dim ctrls As Collection
    Dim item As Variant
    Dim root As Object
    Const q = """"
    Dim errMsg As String
    
    If UCase$(Trim$(zOrderType)) <> "INTERNAL" And UCase$(Trim$(zOrderType)) <> "VISUAL" Then
        errMsg = "[zOrderType] Invalid value: " & q & zOrderType & q & vbLf & "Supported values are " & q & Join(Array("Internal", "Visual"), q & ", " & q) & q
        Err.Raise Number:=516, Description:=errMsg
    End If
    
    Set root = GetUserFormObjectFromCtrl(ctrl)
    
    If ctrl Is root Then
        GetCtrlZOrderIndex = 0
        Exit Function
    End If
    
    Set ctrls = GetAllCtrlsInZOrder(root, zOrderType)
    key = Hex(ObjPtr(ctrl))
    On Error Resume Next
    result = ctrls(key)(2)
    On Error GoTo 0
    
    ' Fall back to linear search using the `Is` operator if key retrieval fails.
    ' This fallback is necessary because Page controls (inside a MultiPage) may yield different
    ' ObjPtr values depending on whether they are typed as Object or MSForms.Control/Page.
    If result = -1 Then
        For Each item In ctrls
            If item(0) Is ctrl Then
                result = item(2)
                Exit For
            End If
        Next item
    End If
    GetCtrlZOrderIndex = result
End Function

Public Function GetCtrlZOrderHierarchy(ByVal ctrl As Object, Optional ByVal zOrderType As String = "Internal") As String
    ' Retrieves the hierarchical Z-Order path string for a target control within a UserForm.
    '
    ' @param ctrl [In] The target control or UserForm object to query.
    ' @param zOrderType  [In] Optional mode to determine how Z-Order is evaluated for Page controls inside a MultiPage.
    '                         Default is "Internal".
    '                           - "Internal": Prioritizes internal Z-Order retrieval for Page controls inside a MultiPage.
    '                                         Controls are retrieved in addition order, even if the visual page order has been moved.
    '                           - "Visual":   Prioritizes visual tab order, assigning Z-Order sequentially starting from the left tab.
    '                                         Reflects any reordering of pages.
    ' @return String A hyphen-delimited string representing the full Z-Order path from top-level container
    '                down to the target control (e.g., "1-3-2").
    '                Returns "0" if the provided object is the top-level UserForm itself.
    '                Returns an empty string ("") if the control is not found.
    '
    ' @raises 516 Thrown if an invalid or unsupported zOrderType string is specified.
    Dim result As String
    Dim key As String
    result = ""
    Dim ctrls As Collection
    Dim item As Variant
    Dim root As Object
    Const q = """"
    Dim errMsg As String
    
    If UCase$(Trim$(zOrderType)) <> "INTERNAL" And UCase$(Trim$(zOrderType)) <> "VISUAL" Then
        errMsg = "[zOrderType] Invalid value: " & q & zOrderType & q & vbLf & "Supported values are " & q & Join(Array("Internal", "Visual"), q & ", " & q) & q
        Err.Raise Number:=516, Description:=errMsg
    End If
    
    Set root = GetUserFormObjectFromCtrl(ctrl)
    
    If ctrl Is root Then
        GetCtrlZOrderHierarchy = "0"
        Exit Function
    End If
    
    Set ctrls = GetAllCtrlsInZOrder(root, zOrderType)
    key = Hex(ObjPtr(ctrl))
    On Error Resume Next
    result = ctrls(key)(3)
    On Error GoTo 0
    
    ' Fall back to linear search using the `Is` operator if key retrieval fails.
    ' This fallback is necessary because Page controls (inside a MultiPage) may yield different
    ' ObjPtr values depending on whether they are typed as Object or MSForms.Control/Page.
    If result = "" Then
        For Each item In ctrls
            If item(0) Is ctrl Then
                result = item(3)
                Exit For
            End If
        Next item
    End If
    GetCtrlZOrderHierarchy = result
End Function

Private Function GetDirectChildCtrls(ByVal parentCtrl As Object) As Collection
    Dim results As Collection
    Set results = New Collection
    Dim ctrl As Object
    Dim root As Object
    
    Set root = GetUserFormObjectFromCtrl(parentCtrl)
    
    If TypeName(parentCtrl) = "MultiPage" Then
        For Each ctrl In parentCtrl.Pages
            results.Add ctrl
        Next
    Else
        For Each ctrl In root.Controls
            If parentCtrl Is ctrl.Parent Then
                results.Add ctrl
            End If
        Next ctrl
    End If
    Set GetDirectChildCtrls = results
End Function

Private Function GetUserFormObjectFromCtrl(ByVal ctrl As Object) As Object
    ' Get the ancestor (UserForm) of the control.
    
    Dim root As Object
    
    If ctrl Is Nothing Then
        Err.Raise 13
    End If
    
    Set root = ctrl
    ' Loop to get root(UserForm) object
    On Error GoTo Finally:
    Do While True
        Set root = root.Parent
    Loop
    On Error GoTo 0

Finally:
    Set GetUserFormObjectFromCtrl = root
End Function

Public Function IsContainerCtrl(ByVal ctrl As Object) As Boolean
    ''
    ' Determines whether the specified control is a container control (e.g., Frame, MultiPage, Page, UserForm)
    ' by checking if it implements the IOleContainer COM interface.
    '
    ' @param ctrl [In] The target control or UserForm object to check.
    ' @return Boolean True if the object implements IOleContainer; otherwise, False.
    ''
    IsContainerCtrl = False
    If ctrl Is Nothing Then Exit Function
    
    Dim pUnk As LongPtr
    pUnk = ObjPtr(ctrl)
    If pUnk = 0 Then Exit Function
    
    ' IID_IOleContainer: {0000011B-0000-0000-C000-000000000046}
    Dim iidContainer(3) As Long
    iidContainer(0) = &H11B: iidContainer(1) = 0: iidContainer(2) = &HC0: iidContainer(3) = &H46000000
    
    Dim pContainer As LongPtr
    Dim hr As Long
    
    ' QueryInterface (QI: VTable Index 0) for IOleContainer
    hr = dcf(pUnk, 0, "QI", VarPtr(iidContainer(0)), VarPtr(pContainer))
    
    If hr = 0 And pContainer <> 0 Then
        IsContainerCtrl = True
        ' Release the retrieved interface (Release: VTable Index 2)
        Call dcf(pContainer, 2, "Release")
    End If
End Function

' IID_IOleContainer: {0000011B-0000-0000-C000-000000000046}
' IID_IEnumUnknown:  {00000100-0000-0000-C000-000000000046}
Private Sub TraceZOrderRecursive(ByVal objContainer As Object, ByVal depth As Long, ByVal currentHierarchy As String, ByRef resultColl As Collection, ByVal useVisualOrder As Boolean)
    ' Internal recursive subroutine that enumerates child controls of a container via COM interfaces
    ' (IOleContainer and IEnumObjects) to accurately determine Z-Order precedence.
    '
    ' @param objContainer      [In]     The container object (UserForm, Frame, MultiPage, or Page).
    ' @param depth             [In]     Current recursion depth level (1-based).
    ' @param currentHierarchy  [In]     Hierarchical Z-Order string of the parent container (e.g., "1-2").
    ' @param resultColl        [In/Out] Target Collection to populate with control metadata.
    ' @param useVisualOrder    [In]     If True, enumerates Page controls within a MultiPage based on their visual
    '                                   tab position rather than internal COM Z-Order.
    Dim pUnk As LongPtr: pUnk = ObjPtr(objContainer)
    Dim pContainer As LongPtr, pEnum As LongPtr
    Dim hr As Long
    Dim children As Collection
    Dim foundType As String
    Dim zOrderIndex As Long
    Dim zOrderHierarchy As String
    
    ' 1. Get IOleContainer
    Dim iidContainer(3) As Long
    iidContainer(0) = &H11B: iidContainer(1) = 0: iidContainer(2) = &HC0: iidContainer(3) = &H46000000
    hr = dcf(pUnk, 0, "QI", VarPtr(iidContainer(0)), VarPtr(pContainer))
    If hr <> 0 Then Exit Sub

    ' 2. EnumObjects
    hr = dcf(pContainer, 4, "EnumObjects", 3, VarPtr(pEnum))
    
    If hr = 0 And pEnum <> 0 Then
        Dim cnt As Long
        
        ' MultiPage -> Get direct Page controls,
        ' Other Containers(e.g., UserForm/Frame/Page) -> Get direct child controls (only controls whose .Parent matches),
        Set children = GetDirectChildCtrls(objContainer)
        cnt = children.Count
        
        If cnt = 0 Then GoTo CleanUp ' Exit if empty
        
        Dim pElements() As LongPtr: ReDim pElements(cnt) ' Allocate with extra capacity
        Dim fetchedCount As Long
        
        Call dcf(pEnum, 5, "Reset")
        
        ' Fetch all specified items at once
        hr = dcf(pEnum, 3, "Next", cnt + 1, VarPtr(pElements(0)), VarPtr(fetchedCount))


        Dim i As Long
        Dim validCount As Long ' Tracks 1-based Z-Order index among matched child controls only (skips hidden internal COM objects)
        ' Calculate 1-based Z-Order index using validCount instead of array index 'i'.
        ' This avoids skipping indices (e.g., "2-2" instead of "2-1") when IOleContainer
        ' returns internal non-control COM objects at the beginning (such as in MultiPage).
        validCount = 0
        
        Dim c As Object
        ' Always declare variable c as Object or Variant.
        ' Declaring it as MsForms.Control may cause some properties of the retrieved control to behave incorrectly.
        ' For example, the Caption property of a Page control is retrieved correctly when c is declared as Object,
        ' but returns an empty string when c is declared as MsForms.Control.
        
        If TypeName(objContainer) = "MultiPage" And useVisualOrder Then
            For Each c In children
                validCount = validCount + 1
                zOrderIndex = validCount
                If currentHierarchy = "" Then
                    zOrderHierarchy = CStr(zOrderIndex)
                Else
                    zOrderHierarchy = currentHierarchy & "-" & CStr(zOrderIndex)
                End If
                resultColl.Add VBA.Array(c, c.Name, zOrderIndex, zOrderHierarchy, depth), Hex(ObjPtr(c))
            
                If IsContainerCtrl(c) Then
                    TraceZOrderRecursive c, depth + 1, zOrderHierarchy, resultColl, useVisualOrder
                End If
            
            Next c
                        
        Else
            For i = 0 To fetchedCount - 1
                foundType = ""
                If pElements(i) <> 0 Then
                    Dim foundName As String: foundName = "Unknown"
                    Dim foundObj As Object: Set foundObj = Nothing
                    
                    ' Normalize IUnknown (for comparison)
                    Dim pUnkReal As LongPtr
                    Dim iidUnk(3) As Long: iidUnk(0) = &H0: iidUnk(1) = 0: iidUnk(2) = &HC0: iidUnk(3) = &H46000000
                    Call dcf(pElements(i), 0, "QI", VarPtr(iidUnk(0)), VarPtr(pUnkReal))
                    
                    ' Match with controls in the current container
    
                    For Each c In children
                        Dim pUnkCtrl As LongPtr
                        Call dcf(ObjPtr(c), 0, "QI", VarPtr(iidUnk(0)), VarPtr(pUnkCtrl))
                        
                        If pUnkReal = pUnkCtrl Then
                            foundName = c.Name
                            foundType = TypeName(c)
                            Set foundObj = c
                            Call dcf(pUnkCtrl, 2, "Release")
                            Exit For
                        End If
                        Call dcf(pUnkCtrl, 2, "Release")
                    Next
                    
                    If foundType <> "" Then
                        validCount = validCount + 1 ' Increment index only when a valid child control is matched
                        zOrderIndex = validCount
                        If currentHierarchy = "" Then
                            zOrderHierarchy = CStr(zOrderIndex)
                        Else
                            zOrderHierarchy = currentHierarchy & "-" & CStr(zOrderIndex)
                        End If
                        resultColl.Add VBA.Array(foundObj, foundName, zOrderIndex, zOrderHierarchy, depth), Hex(ObjPtr(foundObj))
                    End If
                    
                    ' If the found item is a container control(Frame/MultiPage/Page/UserForm), scan inside it recursively
                    If Not foundObj Is Nothing Then
                        If IsContainerCtrl(foundObj) Then
                            TraceZOrderRecursive foundObj, depth + 1, zOrderHierarchy, resultColl, useVisualOrder
                        End If
                    End If
                    
                    ' Cleanup
                    Call dcf(pUnkReal, 2, "Release")
                    Call dcf(pElements(i), 2, "Release")
                End If
            Next
        End If
    End If

CleanUp:
    If pEnum <> 0 Then Call dcf(pEnum, 2, "Release")
    If pContainer <> 0 Then Call dcf(pContainer, 2, "Release")
End Sub


Private Function dcf(ByVal ptr As LongPtr, ByVal vtblIndex As Long, ByVal funcName As String, ParamArray args() As Variant) As Long
    ' Helper wrapper around the Windows API 'DispCallFunc' to invoke COM VTable functions by index.
    '
    ' Note: Copies ParamArray elements into a heap-allocated dynamic array before invocation
    ' to prevent pointer invalidation during execution.
    '
    ' @param ptr       [In] Memory address of the COM interface pointer (IUnknown/IOleContainer/etc.).
    ' @param vtblIndex [In] 0-based method index in the COM Virtual Method Table (VTable).
    ' @param funcName  [In] Descriptive name of the COM method (used for logging/debugging).
    ' @param args()    [In] Variable parameter list passed to the target COM method.
    ' @return Long HRESULT status code from DispCallFunc or the return value of the COM method.
    
    'Debug.Print "dcf called for " & funcName
    Dim l As Long: l = LBound(args)
    Dim u As Long: u = UBound(args)
    Dim cnt As Long: cnt = u - l + 1
    Dim hr As Long, res As Variant
    Dim args_Type() As Integer
    Dim args_Ptr() As LongPtr
    Dim localVar() As Variant
    Const CC_STDCALL As Long = 4 ' Calling convention for the 3rd argument of DispCallFunc (__stdcall)
    ' IMPORTANT: Do NOT use VarPtr(args(i)) directly.
    ' ParamArray elements are temporary Variants managed by the VBA runtime stack.
    ' Their addresses become invalid by the time DispCallFunc internally reads rgpvarg,
    ' causing the COM method to receive garbage values.
    ' Copying into a heap-allocated dynamic array (localArgs) ensures the Variant
    ' addresses remain stable throughout the DispCallFunc call.
    If cnt > 0 Then
        ReDim args_Type(l To u): ReDim args_Ptr(l To u): ReDim localVar(l To u)
        Dim i As Long
        For i = l To u
            localVar(i) = args(i)
            args_Type(i) = VarType(localVar(i))
            args_Ptr(i) = VarPtr(localVar(i))
            'Debug.Print "args(" & i & ")", "Type:" & args_Type(i), "Ptr:" & Hex(args_Ptr(i)),"Value:" & localVar(i)
        Next
        hr = DispCallFunc(ptr, vtblIndex * LenB(ptr), CC_STDCALL, vbLong, cnt, args_Type(l), args_Ptr(l), res)
    Else
        hr = DispCallFunc(ptr, vtblIndex * LenB(ptr), CC_STDCALL, vbLong, cnt, 0, 0, res)
    End If
    If hr = 0 Then
        If res <> 0 Then
            'Debug.Print funcName & " res:" & res
        End If
        dcf = res
    Else
        'Debug.Print funcName & "  hr:" & hr
        dcf = hr
    End If
End Function

