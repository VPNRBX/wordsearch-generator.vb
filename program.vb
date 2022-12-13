Imports System

Module Program

    Dim Board(14, 14)
    Dim StoredWords(9)
    Dim X_Coordinate, Y_Coordinate, Direction As Integer
    Dim Directions = {"Horizontal - Left To Right", "Horizontal - Right To Left", "Vertical - Down", "Vertical - Up", "Diagonal – Down L To R", "Diagonal – Down R To L", "Diagonal – Up L To R", "Diagonal – Up R To L"}

    Dim Randomiser As System.Random = New System.Random() 'Creates the randomiser

    Sub Update_Board() 'A sub-routine that displays the wordsearch with updated information.
        Console.Clear()
        Console.WriteLine("---------------------------------")
        For Y = 0 To 14 ' Produce the grid 1-15 plots.
            Console.Write($"| ")
            For X = 0 To 14
                Console.Write($"{Board(X, Y)} ") 'Using the values stored in the array Board(Value1, Value2) with Value1 being the X axis and Value 2 being the Y axis, display the current
            Next ' information according to that value.
            Console.WriteLine("|")
        Next
        Console.WriteLine("---------------------------------")
        Console.WriteLine()
        For x = 0 To 9 'Write down all the words that have been made for the wordsearch.
            Console.Write($"{StoredWords(x)} ")
        Next
        Console.WriteLine()
        Console.WriteLine()
    End Sub

    Sub Randomise_Unfilled(ByRef FillWith) 'Fills the entire board that is classified as "empty" with the reference value "FillWith"
        For Y = 0 To 14
            For X = 0 To 14
                If FillWith = "-" Then 'Default value.
                    If String.IsNullOrEmpty(Board(X, Y)) Then 'Check if the value in the board in the X and Y axis accordingingly is either Null or empty.
                        Board(X, Y) = "-" 'Make it equal to "-"
                    End If
                Else
                    If String.IsNullOrEmpty(Board(X, Y)) Or Board(X, Y) = "-" Then 'Check if the value in the board in the X and Y axis accordingingly is either Null, empty or set as the
                        'default value "-".
                        Board(X, Y) = Chr(Randomiser.Next(65, 90)) 'Make it equal to a random ASCII character between 65 and 90 (capital letters) and convert it into a character.
                    End If
                End If
            Next
        Next
    End Sub

    Function EnterSubmission(ByRef Word, ByVal Direction, ByVal X_Coordinate, ByVal Y_Coordinate) 'Create a function that returns a true or false value if the spaces for the allocated word
        Dim inc As Integer 'is already occupied.
        If Direction = 0 Then 'Left to Right
            For i = 1 To Len(Word) 'For each character in the referenced "word"
                If String.IsNullOrEmpty(Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1)) Or Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1) = "-" Then 'Check that there is space for
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1) = Mid(Word, i, 1) 'the letter and then make it equate to that character it is currently iterating through.
                    inc += 1 'Increase increment by one.
                Else
                    Return False 'Space is already occupied
                End If
            Next
        End If
        If Direction = 1 Then 'Right to Left
            Word = StrReverse(Word) 'Reverses the string. Additionally you can just do - inc and use the x_coordinate + Len(word) to create the same effect.
            For i = 1 To Len(Word)
                If String.IsNullOrEmpty(Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1)) Or Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1) = "-" Then
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1) = Mid(Word, i, 1)
                    inc += 1
                Else
                    Return False
                End If
            Next
            Word = StrReverse(Word)
        End If
        If Direction = 2 Then 'Up to Down
            For i = 1 To Len(Word)
                If String.IsNullOrEmpty(Board((X_Coordinate) - 1, (Y_Coordinate + inc) - 1)) Or Board((X_Coordinate) - 1, (Y_Coordinate + inc) - 1) = "-" Then
                    Board((X_Coordinate) - 1, (Y_Coordinate + inc) - 1) = Mid(Word, i, 1) 'Increment in y-axis
                    inc += 1
                Else
                    Return False
                End If
            Next
        End If
        If Direction = 3 Then 'Down to Up
            Word = StrReverse(Word)
            For i = 1 To Len(Word)
                If String.IsNullOrEmpty(Board((X_Coordinate) - 1, (Y_Coordinate + inc) - 1)) Or Board((X_Coordinate) - 1, (Y_Coordinate + inc) - 1) = "-" Then
                    Board((X_Coordinate) - 1, (Y_Coordinate + inc) - 1) = Mid(Word, i, 1)
                    inc += 1
                Else
                    Return False
                End If
            Next
            Word = StrReverse(Word)
        End If
        If Direction = 4 Then 'Diagonal - Down L -> R
            For i = 1 To Len(Word)
                If String.IsNullOrEmpty(Board((X_Coordinate + inc) - 1, (Y_Coordinate + inc) - 1)) Or Board((X_Coordinate + inc) - 1, (Y_Coordinate + inc) - 1) = "-" Then
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate + inc) - 1) = Mid(Word, i, 1) 'Increment in y-axis and x-axis.
                    inc += 1
                Else
                    Return False
                End If
            Next
        End If
        If Direction = 5 Then 'Diagonal - Down R -> L
            Word = StrReverse(Word)
            For i = 1 To Len(Word)
                If String.IsNullOrEmpty(Board((X_Coordinate + inc) - 1, (Y_Coordinate + inc) - 1)) Or Board((X_Coordinate + inc) - 1, (Y_Coordinate + inc) - 1) = "-" Then
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate + inc) - 1) = Mid(Word, i, 1)
                    inc += 1
                Else
                    Return False
                End If
            Next
            Word = StrReverse(Word)
        End If
        If Direction = 6 Then 'Diagonal – Up L -> R
            For i = 1 To Len(Word)
                If String.IsNullOrEmpty(Board((X_Coordinate + inc) - 1, (Y_Coordinate - inc) - 1)) Or Board((X_Coordinate + inc) - 1, (Y_Coordinate - inc) - 1) = "-" Then
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate - inc) - 1) = Mid(Word, i, 1) 'Increment in x-axis and decrease in y-axis.
                    inc += 1
                Else
                    Return False
                End If
            Next
        End If
        If Direction = 7 Then 'Diagonal - Up R -> L
            Word = StrReverse(Word)
            For i = 1 To Len(Word)
                If String.IsNullOrEmpty(Board((X_Coordinate + inc) - 1, (Y_Coordinate - inc) - 1)) Or Board((X_Coordinate + inc) - 1, (Y_Coordinate - inc) - 1) = "-" Then
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate - inc) - 1) = Mid(Word, i, 1)
                    inc += 1
                Else
                    Return False
                End If
            Next
            Word = StrReverse(Word)
        End If
        Return True
    End Function

    Sub RequestCoordinates(ByRef Current) 'Sub-routine that gathers the co-ordinates inputted by user.
        Do
            Console.WriteLine("Please enter where you wish the word to start on the grid:") 'Ask the user where they want the word to begin on the grid.
            Console.Write("X: ")
            X_Coordinate = Console.ReadLine() 'Assign these variables according to which axis.
            Console.Write("Y: ")
            Y_Coordinate = Console.ReadLine()
            If X_Coordinate < 1 Or Y_Coordinate < 1 Or X_Coordinate > 15 Or Y_Coordinate > 15 Then 'Check that assigned values are within range of the board.
                Console.Clear()
                Console.WriteLine("Please enter a valid Y and X co-ordinate on the grid that is between points 1 and 15.") 'If not, display error message.
                Threading.Thread.Sleep(1000)
                Console.Clear()
                Update_Board()
                RequestCoordinates(Current) 'Rerun subroutine.
            End If
            If Not EnterSubmission(StoredWords(Current), Direction, X_Coordinate, Y_Coordinate) Then 'Await return function for if the submitted word is in a valid spot or not (pre-occupied).
                Console.Clear() '(Returned false).
                Console.WriteLine("Please enter a valid Y and X co-ordinate between points 1 and 15 that is not already occupied.") 'Displays error message that it's already occupied.
                Threading.Thread.Sleep(1000)
                Console.Clear()
                Update_Board()
                RequestCoordinates(Current) 'Rerun subroutine.
            End If
        Loop Until X_Coordinate <= 15 And Y_Coordinate <= 15 And X_Coordinate >= 1 And Y_Coordinate >= 1 'Loop until the variables are valid and between the values.
    End Sub

    Sub Main(args As String())
        Randomise_Unfilled("-") 'Set all plots in the grid to "-"
        For Current = 0 To 9 'Repeat 10 times (10 words).
            Console.Clear()
            Update_Board()
            Console.WriteLine("Please enter a word:")
            StoredWords(Current) = UCase(Console.ReadLine()) 'Store the word they inputed as capital letters in the array StoredWords.
            Console.WriteLine()
            Dim Selected As Integer
            Do
                Console.WriteLine("Please enter a direction from the following: ")
                For List = 0 To 7
                    Console.WriteLine($"{List}. {Directions(List)}") 'Print out the directions they can pick.
                Next
                Selected = Console.ReadLine()
                If Selected > 7 Or Selected < 0 Then 'Check selected is within the valid bounds.
                    Console.Clear()
                    Console.WriteLine("Please enter a valid number.") 'Print error if its not.
                    Threading.Thread.Sleep(1000)
                    Console.Clear()
                End If
            Loop Until Selected <= 7 And Selected >= 0 'Loop until selected is in valid parameters.
            Direction = Selected
            Console.WriteLine()
            RequestCoordinates(Current) 'Run the requested subroutine with the current word.
        Next
        Randomise_Unfilled("True") 'After 10 words have been inputted and stored into the board, change every default plot to a randomise character.
        Update_Board() 'Update the board with the final word search.

    End Sub
End Module
