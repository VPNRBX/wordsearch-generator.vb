Imports System

Module Program

    Dim Board(14, 14)
    Dim StoredWords(9)
    Dim X_Coordinate, Y_Coordinate, Direction As Integer
    Dim Directions = {"Horizontal - Left To Right", "Horizontal - Right To Left", "Vertical - Down", "Vertical - Up", "Diagonal – Down L To R", "Diagonal – Down R To L", "Diagonal – Up L To R", "Diagonal – Up R To L"}

    Dim Randomiser As System.Random = New System.Random()

    Sub Update_Board()
        Console.Clear()
        Console.WriteLine("---------------------------------")
        For Y = 0 To 14
            Console.Write($"| ")
            For X = 0 To 14
                Console.Write($"{Board(X, Y)} ")
            Next
            Console.WriteLine("|")
        Next
        Console.WriteLine("---------------------------------")
        Console.WriteLine()
        For x = 0 To 9
            Console.Write($"{StoredWords(x)} ")
        Next
        Console.WriteLine()
        Console.WriteLine()
    End Sub

    Sub Randomise_Unfilled(ByRef FillWith)
        For Y = 0 To 14
            For X = 0 To 14
                If FillWith = "-" Then
                    If String.IsNullOrEmpty(Board(X, Y)) Then
                        Board(X, Y) = "-"
                    End If
                Else
                    If String.IsNullOrEmpty(Board(X, Y)) Or Board(X, Y) = "-" Then
                        Board(X, Y) = Chr(Randomiser.Next(65, 90))
                    End If
                End If
            Next
        Next
    End Sub

    Function EnterSubmission(ByRef Word, ByVal Direction, ByVal X_Coordinate, ByVal Y_Coordinate)
        Dim inc As Integer
        If Direction = 0 Then 'Left to Right
            For i = 1 To Len(Word)
                If String.IsNullOrEmpty(Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1)) Or Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1) = "-" Then
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate) - 1) = Mid(Word, i, 1)
                    inc += 1
                Else
                    Return False
                End If
            Next
        End If
        If Direction = 1 Then 'Right to Left
            Word = StrReverse(Word)
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
                    Board((X_Coordinate) - 1, (Y_Coordinate + inc) - 1) = Mid(Word, i, 1)
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
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate + inc) - 1) = Mid(Word, i, 1)
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
                    Board((X_Coordinate + inc) - 1, (Y_Coordinate - inc) - 1) = Mid(Word, i, 1)
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

    Sub RequestCoordinates(ByRef Current)
        Do
            Console.WriteLine("Please enter where you wish the word to start on the grid:")
            Console.Write("X: ")
            X_Coordinate = Console.ReadLine()
            Console.Write("Y: ")
            Y_Coordinate = Console.ReadLine()
            If X_Coordinate < 1 Or Y_Coordinate < 1 Or X_Coordinate > 15 Or Y_Coordinate > 15 Then
                Console.Clear()
                Console.WriteLine("Please enter a valid Y and X co-ordinate on the grid that is between points 1 and 15.")
                Threading.Thread.Sleep(1000)
                Console.Clear()
                Update_Board()
                RequestCoordinates(Current)
            End If
            If Not EnterSubmission(StoredWords(Current), Direction, X_Coordinate, Y_Coordinate) Then
                Console.Clear()
                Console.WriteLine("Please enter a valid Y and X co-ordinate between points 1 and 15 that is not already occupied.")
                Threading.Thread.Sleep(1000)
                Console.Clear()
                Update_Board()
                RequestCoordinates(Current)
            End If
        Loop Until X_Coordinate <= 15 And Y_Coordinate <= 15 And X_Coordinate >= 1 And Y_Coordinate >= 1
    End Sub

    Sub Main(args As String())
        Randomise_Unfilled("-")
        For Current = 0 To 9
            Console.Clear()
            Update_Board()
            Console.WriteLine("Please enter a word:")
            StoredWords(Current) = UCase(Console.ReadLine())
            Console.WriteLine()
            Dim Selected As Integer
            Do
                Console.WriteLine("Please enter a direction from the following: ")
                For List = 0 To 7
                    Console.WriteLine($"{List}. {Directions(List)}")
                Next
                Selected = Console.ReadLine()
                If Selected > 7 Or Selected < 0 Then
                    Console.Clear()
                    Console.WriteLine("Please enter a valid number.")
                    Threading.Thread.Sleep(1000)
                    Console.Clear()
                End If
            Loop Until Selected <= 7 And Selected >= 0
            Direction = Selected
            Console.WriteLine()
            RequestCoordinates(Current)
        Next
        Randomise_Unfilled("True")
        Update_Board()

    End Sub
End Module
