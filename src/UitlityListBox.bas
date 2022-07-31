Attribute VB_Name = "UitlityListBox"
Option Explicit
Option Compare Database

' http://allenbrowne.com/func-12.html

Public Function ListBoxClearList(ByVal lst As ListBox) As Boolean
    'Purpose:   Unselect all items in the listbox.
    'Return:    True if successful
    'Author:    Allen Browne. http://allenbrowne.com  June, 2006.
    Dim varItem As Variant

    If lst.MultiSelect = 0 Then
        lst = Null
    Else
        For Each varItem In lst.ItemsSelected
            lst.Selected(varItem) = False
        Next
    End If

    ListBoxClearList = True

End Function

Public Function ListBoxSelectAll(ByVal lst As ListBox) As Boolean

    'Purpose:   Select all items in the multi-select list box.
    'Return:    True if successful
    'Author:    Allen Browne. http://allenbrowne.com  June, 2006.
    Dim lngRow As Long

    If lst.MultiSelect Then
        For lngRow = 0 To lst.ListCount - 1
            lst.Selected(lngRow) = True
        Next
        ListBoxSelectAll = True
    End If

ListBoxSelectAll = True
End Function
