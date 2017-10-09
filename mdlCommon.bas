Attribute VB_Name = "mdlCommon"

Public Sub CopyListview(lvSource As ListView, lvDest As ListView)

    Dim iSubCnt As Integer 'Holds the subItems
    Dim iList As ListItem
    Dim OList As ListItem


Exit Sub

    Set iList = lvSource.SelectedItem 'Copy from the first list view
    iSubCnt = lvSource.ColumnHeaders.count - 1 '-1 for the listitem

    Set OList = lvDest.ListItems.Add(, iList.Key, iList.Text)

    With lvSource.SelectedItem
        If iSubCnt >= 1 Then 'then we have subitems
            For intLoop = 0 To iSubCnt - 1
                OList.SubItems(intLoop + 1) = .SubItems(intLoop + 1)
            Next
        End If
    End With
'
'    'lvSource.ListItems.Remove lvSource.SelectedItem.Key
End Sub
