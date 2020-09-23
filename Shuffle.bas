Attribute VB_Name = "Shuffle"
Public Card(52) As Integer
Dim CardValue, x As Integer

'This module is currently set up to shuffle only a
'single deck of 52 cards.
'The cards, when shuffled, are stored as random
'integer values from 1 to 52 in the Card() array.
'Card values are not duplicated within the array.

Public Sub ShuffleDeck()
x = 0
Erase Card
'This is where the random shuffle is located.
Randomize
CardValue = Int((52 * Rnd) + 1)
Card(0) = CardValue
For x = 1 To 51
    Do While IsInArray = True
        Randomize
        CardValue = Int((52 * Rnd) + 1)
        IsInArray
    Loop
Card(x) = CardValue
Next x
End Sub

Public Function IsInArray() As Boolean
'Compares the next random number against the numbers
'currently stored in the Card() array. If it is a
'duplicate, the random number is discarded and is
'generated again, then tested again until it is found
'to be unique within the array.
Dim y As Integer
For y = 0 To x
    If CardValue = Card(y) Then
        IsInArray = True
        Exit Function
    End If
Next y
IsInArray = False
End Function


