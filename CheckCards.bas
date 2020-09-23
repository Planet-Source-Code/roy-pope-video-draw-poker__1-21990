Attribute VB_Name = "CheckCards"
Public SortVariable(5) As Integer

Public Sub BubbleSort()
'Sort the array of the cards dealt, after the cards have been reduced, in
'ascending order. This allows for computation of the straights more than
'anything else.
Dim n, i, j, s As Integer
    n = 4
    For i = 1 To n + 1
        For j = 0 To n - i
            If SortVariable(j) > SortVariable(j + 1) Then
                s = SortVariable(j)
                SortVariable(j) = SortVariable(j + 1)
                SortVariable(j + 1) = s
            End If
        Next j
    Next i
End Sub

Public Sub ReduceCards()
'Reduces the number value of the cards dealt so that all cards fall into the
'range between 1 nd 52.
    For x = 0 To 4
        Select Case SortVariable(x)
            Case 14 To 26: SortVariable(x) = SortVariable(x) - 13
            Case 27 To 39: SortVariable(x) = SortVariable(x) - (13 * 2)
            Case 40 To 52: SortVariable(x) = SortVariable(x) - (13 * 3)
        End Select
    Next x
End Sub

Public Function DoIsStraight() As Boolean
'Checks the cards for a straight in the hand. This function only checks for
'straights in which the ace is the small-card.
Dim Straight As Boolean
If SortVariable(1) = SortVariable(0) + 1 Then
    If SortVariable(2) = SortVariable(1) + 1 Then
        If SortVariable(3) = SortVariable(2) + 1 Then
            If SortVariable(4) = SortVariable(3) + 1 Then
                DoIsStraight = True
                Exit Function
            End If
        End If
    End If
Else
    DoIsStraight = False
End If
End Function

Public Function DoIsFlush() As Boolean
Dim MinRange, MaxRange As Integer
MinRange = 1
MaxRange = 13
For x = 1 To 4
    If SortVariable(0) >= MinRange And SortVariable(0) <= MaxRange Then
        If SortVariable(1) >= MinRange And SortVariable(1) <= MaxRange Then
            If SortVariable(2) >= MinRange And SortVariable(2) <= MaxRange Then
                If SortVariable(3) >= MinRange And SortVariable(3) <= MaxRange Then
                    If SortVariable(4) >= MinRange And SortVariable(4) <= MaxRange Then
                        DoIsFlush = True
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    MinRange = MinRange + 13
    MaxRange = MaxRange + 13
Next x
DoIsFlush = False
End Function
    
Public Function DoIs4Kind() As Boolean
Select Case SortVariable(0)
    Case SortVariable(1):
        If SortVariable(1) = SortVariable(2) And SortVariable(2) = SortVariable(3) Then DoIs4Kind = True
        If SortVariable(1) = SortVariable(2) And SortVariable(2) = SortVariable(4) Then DoIs4Kind = True
        If SortVariable(1) = SortVariable(3) And SortVariable(3) = SortVariable(4) Then DoIs4Kind = True
    Case SortVariable(2):
        If SortVariable(2) = SortVariable(3) And SortVariable(3) = SortVariable(4) Then DoIs4Kind = True
End Select
If SortVariable(1) = SortVariable(2) And SortVariable(2) = SortVariable(3) And SortVariable(3) = SortVariable(4) Then
    DoIs4Kind = True
End If
End Function

Public Function DoIs3Kind() As Boolean
'Function to check all possible combinations of cards
'drawn to see if there is a three of a kind in the hand.
DoIs3Kind = False 'There is no three of a kind
'All the results of the code which follows indicate that
'a three of a kind is found.
Select Case SortVariable(0)
    Case SortVariable(1):
        If SortVariable(1) = SortVariable(2) Then DoIs3Kind = True
        If SortVariable(1) = SortVariable(3) Then DoIs3Kind = True
        If SortVariable(1) = SortVariable(4) Then DoIs3Kind = True
    Case SortVariable(2):
        If SortVariable(2) = SortVariable(3) Then DoIs3Kind = True
        If SortVariable(2) = SortVariable(4) Then DoIs3Kind = True
    Case SortVariable(3):
        If SortVariable(3) = SortVariable(4) Then DoIs3Kind = True
End Select
Select Case SortVariable(1)
    Case SortVariable(2):
        If SortVariable(2) = SortVariable(3) Then DoIs3Kind = True
        If SortVariable(2) = SortVariable(4) Then DoIs3Kind = True
    Case SortVariable(3):
        If SortVariable(3) = SortVariable(4) Then DoIs3Kind = True
End Select
If SortVariable(2) = SortVariable(3) And SortVariable(3) = SortVariable(4) Then DoIs3Kind = True
End Function

Public Function DoIsPair() As Boolean
'Function to check all possible combinations of cards
'drawn to see if there is a pair in the hand.
DoIsPair = False 'There is no pair in the hand.
'All the results of the code which follows indicate that
'a pair is found.
Select Case SortVariable(0)
    Case SortVariable(1): DoIsPair = True
    Case SortVariable(2): DoIsPair = True
    Case SortVariable(3): DoIsPair = True
    Case SortVariable(4): DoIsPair = True
End Select
Select Case SortVariable(1)
    Case SortVariable(2): DoIsPair = True
    Case SortVariable(3): DoIsPair = True
    Case SortVariable(4): DoIsPair = True
End Select
Select Case SortVariable(2)
    Case SortVariable(3): DoIsPair = True
    Case SortVariable(4): DoIsPair = True
End Select
If SortVariable(3) = SortVariable(4) Then DoIsPair = True
End Function

Public Function DoIsRoyal() As Boolean
DoIsRoyal = False
If SortVariable(0) = 1 Then
    If SortVariable(1) = 10 Then
        If SortVariable(2) = 11 Then
            If SortVariable(3) = 12 Then
                If SortVariable(4) = 13 Then
                    DoIsRoyal = True
                End If
            End If
        End If
    End If
End If
End Function

Public Function DoIsFullHouse() As Boolean
'Function to check all five cards drawn to see if the hand is a
'full house.
DoIsFullHouse = False 'Not a full house.
'All the results of the code which follows indicate that
'a full house is found.
Select Case SortVariable(0)
    Case SortVariable(1):
        If SortVariable(1) = SortVariable(2) And SortVariable(3) = SortVariable(4) Then DoIsFullHouse = True
        If SortVariable(1) = SortVariable(3) And SortVariable(2) = SortVariable(4) Then DoIsFullHouse = True
        If SortVariable(1) = SortVariable(4) And SortVariable(2) = SortVariable(3) Then DoIsFullHouse = True
    Case SortVariable(2):
        If SortVariable(2) = SortVariable(3) And SortVariable(1) = SortVariable(4) Then DoIsFullHouse = True
        If SortVariable(2) = SortVariable(4) And SortVariable(1) = SortVariable(3) Then DoIsFullHouse = True
    Case SortVariable(3):
        If SortVariable(3) = SortVariable(4) And SortVariable(1) = SortVariable(2) Then DoIsFullHouse = True
End Select
Select Case SortVariable(1)
    Case SortVariable(2):
        If SortVariable(2) = SortVariable(3) And SortVariable(0) = SortVariable(4) Then DoIsFullHouse = True
        If SortVariable(2) = SortVariable(4) And SortVariable(0) = SortVariable(3) Then DoIsFullHouse = True
    Case SortVariable(3):
        If SortVariable(3) = SortVariable(4) And SortVariable(0) = SortVariable(2) Then DoIsFullHouse = True
End Select
If SortVariable(2) = SortVariable(3) And SortVariable(3) = SortVariable(4) Then
    If SortVariable(0) = SortVariable(1) Then
        DoIsFullHouse = True
    End If
End If
End Function

Public Function DoIsTwoPair() As Boolean
'This function is called only if a single pair is first found. The function
'examines the cards other than the two originally ofund to be a pair, and
'determines if a second pair is also located within the player's hand.
DoIsTwoPair = False 'No second pair is found.
'All the results of the code which follows indicate that
'a second pair is found.
Select Case SortVariable(0)
    Case SortVariable(1):
        If SortVariable(2) = SortVariable(3) Then DoIsTwoPair = True
        If SortVariable(2) = SortVariable(4) Then DoIsTwoPair = True
        If SortVariable(3) = SortVariable(4) Then DoIsTwoPair = True
    Case SortVariable(2):
        If SortVariable(1) = SortVariable(3) Then DoIsTwoPair = True
        If SortVariable(1) = SortVariable(4) Then DoIsTwoPair = True
        If SortVariable(3) = SortVariable(4) Then DoIsTwoPair = True
    Case SortVariable(3):
        If SortVariable(1) = SortVariable(2) Then DoIsTwoPair = True
        If SortVariable(1) = SortVariable(4) Then DoIsTwoPair = True
        If SortVariable(2) = SortVariable(4) Then DoIsTwoPair = True
    Case SortVariable(4):
        If SortVariable(1) = SortVariable(2) Then DoIsTwoPair = True
        If SortVariable(1) = SortVariable(3) Then DoIsTwoPair = True
        If SortVariable(2) = SortVariable(3) Then DoIsTwoPair = True
End Select
Select Case SortVariable(1)
    Case SortVariable(2):
        If SortVariable(0) = SortVariable(3) Then DoIsTwoPair = True
        If SortVariable(0) = SortVariable(4) Then DoIsTwoPair = True
        If SortVariable(3) = SortVariable(4) Then DoIsTwoPair = True
    Case SortVariable(3):
        If SortVariable(0) = SortVariable(2) Then DoIsTwoPair = True
        If SortVariable(0) = SortVariable(4) Then DoIsTwoPair = True
        If SortVariable(2) = SortVariable(4) Then DoIsTwoPair = True
    Case SortVariable(4):
        If SortVariable(0) = SortVariable(2) Then DoIsTwoPair = True
        If SortVariable(0) = SortVariable(3) Then DoIsTwoPair = True
        If SortVariable(2) = SortVariable(3) Then DoIsTwoPair = True
End Select
Select Case SortVariable(2)
    Case SortVariable(3):
        If SortVariable(0) = SortVariable(1) Then DoIsTwoPair = True
        If SortVariable(0) = SortVariable(4) Then DoIsTwoPair = True
        If SortVariable(1) = SortVariable(4) Then DoIsTwoPair = True
    Case SortVariable(4):
        If SortVariable(0) = SortVariable(1) Then DoIsTwoPair = True
        If SortVariable(0) = SortVariable(3) Then DoIsTwoPair = True
        If SortVariable(1) = SortVariable(3) Then DoIsTwoPair = True
End Select
If SortVariable(3) = SortVariable(4) Then
    If SortVariable(0) = SortVariable(1) Then DoIsTwoPair = True
    If SortVariable(0) = SortVariable(2) Then DoIsTwoPair = True
    If SortVariable(1) = SortVariable(2) Then DoIsTwoPair = True
End If
End Function
