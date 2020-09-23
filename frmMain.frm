VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdScramble 
      Caption         =   "Scramble List"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   4455
   End
   Begin VB.ListBox TheList 
      Height          =   2010
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0013
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.TextBox txtFocusHolder 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblLastIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblFirstIndex 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChoosingFirst As Boolean, ChoosingLast As Boolean, TheNums()

Enum lOrder
    Ascending
    Descending
End Enum

'This will generate an array of non-repeating random numbers 0 to intMax
'The original code for this procedure was written by Kevin Lawrence as a
'button response and rewritten by me as shown below.
Private Sub RandomList(ByVal intMax As Integer)
    Dim High(), Nums(), iMax As Integer
    'Holds the intMax variable minus 1
    iMax% = intMax% - 1
    'An array for holding the highest number to choose
    ReDim High(iMax%)
    'The new array
    ReDim Nums(iMax%)
    'The High array will hold a list of sequential numbers from 0 to intMax - 1.
    'These numbers will be inserted randomly into the new array.
    For i% = 0 To iMax%
        High(i%) = i%
    Next
    'Work backwards from iMax% to 0
    For i% = iMax% To 0 Step -1
        'Choose a random array index
        Chosen% = Int(i% * Rnd)
        'Insert an element from High into the new array
        Nums(iMax% - i%) = High(Chosen%)
        'Replace the chosen element from High with a new number
        High(Chosen%) = High(i%)
    Next
    'Replace existing (or not) TheNums array with the new array
    TheNums = Nums
End Sub

'This will scramble the specified ListBox
Private Sub ScrambleList(ByVal lList As ListBox)
    Dim arrList() As String
    'Resize the array to the size of the list and subtract 1
    'because the index is 0 based
    ReDim arrList(lList.ListCount - 1)
    'Retrieve an array of random numbers to set the order
    'of the list
    RandomList lList.ListCount
    'Randomly insert the items into the arrList array
    For i% = 0 To UBound(arrList)
        arrList(i%) = lList.List(TheNums(i%))
    Next
    'Return the randomized list items to the list
    For i% = 0 To UBound(arrList)
        lList.List(i%) = arrList(i%)
    Next
End Sub

'This will sort the array in ascending order. Don't worry, this won't effect the resulting list.
Public Sub SortArray(arrList())
    Dim strTemp As String, InOrder As Boolean
    'c will hold the index of which characters are being compared
    'l will hold the length of the longest array item between the two currently being compared
    Dim c As Integer, l As Integer
    't1 and t2 will hold the string version of the array items.
    Dim t1 As String, t2 As String
    'c1 and c2 will hold individual uppercase characters for comparing
    Dim c1 As String, c2 As String
    'Loop through the array
    For i% = 0 To UBound(arrList()) - 1
        'Loop through the array items following array item i%
        For j% = i% + 1 To UBound(arrList())
            't1$ and t2$ will hold the string version of the array items.
            t1$ = CStr(arrList(i%))
            t2$ = CStr(arrList(j%))
            'If t1$ is longer than t2$, then
            If Len(t1$) > Len(t2$) Then
                'l% is that length
                l% = Len(t1$)
            Else
                'Otherwise l% is the length of t2$
                l% = Len(t2$)
            End If
            'Loop through the characters of the array items
            For c% = 1 To l%
                'c1$ and c2$ will hold individual uppercase characters for comparing
                c1$ = UCase$(Mid$(t1$, c%, 1))
                c2$ = UCase$(Mid$(t2$, c%, 1))
                'If the characters have been the same up to this point and there
                'are no more characters to loop through in the first array item,
                'then the items are in order.
                If Len(c1$) = 0 Then
                    InOrder = True
                    GoTo skip
                'If there are no more characters to loop through in the second array
                'item, then the items are not in order.
                ElseIf Len(c2$) = 0 Then
                    InOrder = False
                    GoTo skip
                End If
                'If the current character from the first array item comes before the
                'character from the other, then the two items are in order.
                If Asc(c1$) < Asc(c2$) Then
                    InOrder = True
                    GoTo skip
                'If the current character from the first array item comes before the
                'character from the other, then the two items are in order.
                ElseIf Asc(c2$) < Asc(c1$) Then
                    InOrder = False
                    GoTo skip
                End If
                'If the two characters are the same, then keep looping throught the characters
            Next
skip:
            'If the items are not in order, switch them around. Store an array item in a temporary
            'variable while replacing its value with the other array item's value. Then replace that
            'value with the value stored in the temporary string.
            If Not InOrder Then
                strTemp$ = arrList(j%)
                arrList(j%) = arrList(i%)
                arrList(i%) = strTemp$
            End If
        Next
    Next
End Sub

Private Sub SortList(ByVal lList As ListBox, ByVal FirstIndex As Integer, ByVal LastIndex As Integer, ByVal Order As lOrder)
    Dim arrList(), NewFirstIndex As Integer, NewLastIndex As Integer, NewStep As Integer, OtherIndex As Integer
    'Check for obvious errors
    ' - If FirstIndex comes at or after LastIndex. It cannot be equal to LastIndex
    '   because then you're only asking to sort one item.
    If FirstIndex% >= LastIndex% Then
        MsgBox "FirstIndex must be less than LastIndex."
        Exit Sub
    End If
    ' - If the list has only one or no items, then you don't even need to sort.
    If lList.ListCount < 2 Then
        MsgBox lList.Name & " does not contain enough list items."
        Exit Sub
    '   otherwise
    Else
        ' - If the list only has two items, then FirstIndex has to be 0.
        If lList.ListCount = 2 And FirstIndex% <> 0 Then
            MsgBox "FirstIndex must be equal to 0."
            Exit Sub
        ' - otherwise if FirstIndex is < 0 or > the second to last item, then it is invalid
        ElseIf FirstIndex% < 0 Or FirstIndex% > lList.ListCount - 2 Then
            MsgBox "FirstIndex must be between or equal to 0 and " & lList.ListCount - 2 & "."
            Exit Sub
        End If
        ' - If the list only has two items, then LastIndex must be 1.
        If lList.ListCount = 2 And LastIndex% <> 1 Then
            MsgBox "LastIndex must be equal to 1."
            Exit Sub
        ' - otherwise if LastIndex < 1 or > the number of items in the list minus 1, then it is invalid
        ElseIf LastIndex% < 1 Or LastIndex% > lList.ListCount - 1 Then
            MsgBox "LastIndex must be between or equal to 1 and " & lList.ListCount - 1 & "."
            Exit Sub
        End If
    End If
    'Resize the arrList array to the fit the number of chosen list items
    ReDim arrList(LastIndex% - FirstIndex%)
    'Fill the arrList array with the chosen list items
    For i% = FirstIndex% To LastIndex%
        arrList(i% - FirstIndex%) = lList.List(i%)
    Next
    'Sort the arrList array in ascending order
    SortArray arrList()
    'Specify the direction to change the existing list items
    If Order = Descending Then
        NewFirstIndex% = LastIndex%
        NewLastIndex% = FirstIndex%
        NewStep% = -1
    Else
        NewFirstIndex% = FirstIndex%
        NewLastIndex% = LastIndex%
        NewStep% = 1
    End If
    'The index for the arrList array
    OtherIndex% = 0
    'Replace the existing list items with the new sorted list items
    For i = NewFirstIndex% To NewLastIndex% Step NewStep%
        lList.List(i) = arrList(OtherIndex%)
        OtherIndex% = OtherIndex% + 1
    Next
End Sub

Private Sub cmdScramble_Click()
    ScrambleList TheList
End Sub

Private Sub cmdSort_Click()
    'You may choose ascending or descending
    SortList TheList, Val(lblFirstIndex.Caption) \ 1, Val(lblLastIndex.Caption) \ 1, Ascending
End Sub

Private Sub Form_Load()
    Dim c As Control
    ScrambleList TheList
    For Each c In Me.Controls
        If TypeOf c Is Label Then c.Top = c.Top + 15
    Next
    lblFirstIndex.Caption = 0
    lblLastIndex.Caption = TheList.ListCount - 1
    Me.Show
    txtFocusHolder.SetFocus
End Sub

Private Sub TheList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim li As Integer
    
    li = TheList.ListIndex
    TheList.ListIndex = -1
    
    txtFocusHolder.SetFocus
    
    If Button = 1 Then
        If ChoosingFirst Then
            lblFirstIndex.Caption = li
            ChoosingFirst = False
        ElseIf ChoosingLast Then
            lblLastIndex.Caption = li
            ChoosingLast = False
        End If
    End If
End Sub

Private Sub lblFirstIndex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ChoosingLast = False
        ChoosingFirst = (Not ChoosingFirst)
        If ChoosingFirst Then MsgBox "Choose an item from the list as the first list item in the list to be sorted."
    End If
End Sub

Private Sub lblLastIndex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ChoosingFirst = False
        ChoosingLast = (Not ChoosingLast)
        If ChoosingLast Then MsgBox "Choose an item from the list as the last list item in the list to be sorted."
    End If
End Sub
