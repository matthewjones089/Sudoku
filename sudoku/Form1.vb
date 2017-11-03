Public Class Form1

    ' Cell Structure
    ' This holds the current cell value, whether it is valid or not, and details about the button object

    Public Structure cell
        Public value As Integer
        Public valid As Boolean
        Public button As Button
    End Structure

    ' Create a 9 x 9 array to hold the cell information and button objects

    Dim c(8, 8) As cell

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim x As Integer
        Dim y As Integer
        Dim x_offset As Integer
        Dim y_offset As Integer

        ' Gives the form a title
        Me.Text = "Sudoku Challenge"

        ' When the form loads we need to initialize the cell structure and create the buttons

        ' Process each row

        For y = 0 To 8

            ' Set the default y-offset of the button to 10 pixels

            y_offset = 10

            ' If the cell y-number is 3, 4 or 5 then increment the y-offset by 10 pixels

            If y >= 3 And y <= 5 Then
                y_offset = y_offset + 10
            Else

                ' If the cell y-number is 6, 7, or 8 then increment the y-offset by 20 pixels

                If y >= 6 Then
                    y_offset = y_offset + 20
                End If

            End If

            ' Process each column

            For x = 0 To 8

                ' Set the default x-offset of the button to 10pixels

                x_offset = 10

                ' If the cell x-number is 3, 4 or 5 then increment the x-offset by 10 pixels

                If x >= 3 And x <= 5 Then
                    x_offset = x_offset + 10
                Else

                    ' If the cell x-number is 6, 7, or 8 then increment the x-offset by 20 pixels

                    If x >= 6 Then
                        x_offset = x_offset + 20
                    End If

                End If

                ' Create a new button object in the cell structure

                c(x, y).button = New Button

                ' Set the text on the button to blank

                c(x, y).button.Text = ""

                ' Set the x and y position of the button

                c(x, y).button.Left = (x * 50) + x_offset
                c(x, y).button.Top = (y * 50) + y_offset

                ' Set the height and width of the button

                c(x, y).button.Width = 45
                c(x, y).button.Height = 45

                ' Store the number of button in the button tag

                c(x, y).button.Tag = (y * 9) + x

                ' Set the font size of the button to 26 and set it to bold

                c(x, y).button.Font = New Font(c(x, y).button.Font.FontFamily, 26, FontStyle.Bold)

                ' Set each button background colour to white

                c(x, y).button.BackColor = Color.White

                ' Set the click event of the button to call function button_click

                AddHandler c(x, y).button.Click, AddressOf button_click

                ' Add the button to the form

                Me.Controls.Add(c(x, y).button)

            Next
        Next

        create_game()

        ' Set the form height and width

        Me.Height = 485
        Me.Width = 485

        ' Set form as fixed size

        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle

    End Sub

    ' This function is called when a button is clicked

    Sub button_click(ByVal sender As Object, ByVal e As EventArgs)

        Dim btn As Button
        Dim x As Integer
        Dim y As Integer
        Dim number As Integer
        Dim all_valid As Boolean

        ' Get the current button object

        btn = sender

        ' Work out its x and y position by using the button number stored in the button tag

        y = Int(btn.Tag / 9)
        x = btn.Tag - (y * 9)

        ' Prompt for a valid number

        number = get_number()

        ' If the number is zero then reset the cell

        If (number = 0) Then
            c(x, y).button.Text = ""
        Else

            ' If a number has been entered then set the button text to the number

            c(x, y).button.Text = number

        End If

        ' Set the value of the current cell to the number entered

        c(x, y).value = number

        ' Set all cells to valid initially

        For y = 0 To 8
            For x = 0 To 8
                c(x, y).valid = True
            Next
        Next

        ' Check each cell to determine whether it is invalid

        For y = 0 To 8
            For x = 0 To 8
                check_cell(x, y)
            Next
        Next

        ' Set the colour of each button:
        ' If the button has no value then it is white
        ' If the button has a value but is invalid then it is red
        ' If the button has a value and is valid then it is light green

        For y = 0 To 8
            For x = 0 To 8
                If c(x, y).value = 0 Then
                    c(x, y).button.BackColor = Color.White
                Else
                    If c(x, y).valid = True Then
                        c(x, y).button.BackColor = Color.LightGreen
                    Else
                        c(x, y).button.BackColor = Color.Red
                    End If
                End If
            Next
        Next

        ' Check whether the puzzle is complete
        ' All cells are valid initially, and then we check whether any are not valid

        all_valid = True

        ' Check each cell

        For y = 0 To 8
            For x = 0 To 8

                ' If a cell has no value or if it is invalid then the puzzle is incomplete

                If (c(x, y).valid = False Or c(x, y).value = 0) Then
                    all_valid = False
                End If

            Next
        Next

        ' If the puzzle is complete and all cells are valid then show a message 

        If all_valid = True Then
            MsgBox("Well done you solved the puzzle!")
        End If

    End Sub

    Function check_cell(ByVal x As Integer, ByVal y As Integer) As Boolean

        ' If the cell has a value then we should check whether it is valid

        If c(x, y).value <> 0 Then
            check_row(x, y)
            check_col(x, y)
            check_3by3(x, y)
        End If

        Return True

    End Function

    Function check_row(ByVal x As Integer, ByVal y As Integer) As Boolean

        Dim x1 As Integer
        Dim value As Integer

        ' Get the value of the cell we are checking

        value = c(x, y).value

        ' Check each cell in the current row

        For x1 = 0 To 8

            ' We should ignore the cell we are checking

            If (x1 <> x) Then

                ' If any cell in the row has the same value as the one we are checking then
                ' mark it as invalid

                If (c(x1, y).value = value) Then
                    c(x1, y).valid = False
                End If
            End If
        Next

        Return True

    End Function

    Function check_col(ByVal x As Integer, ByVal y As Integer) As Boolean

        Dim y1 As Integer
        Dim value As Integer

        ' Get the value of the cell we are checking

        value = c(x, y).value

        ' Check each cell in the current column

        For y1 = 0 To 8

            ' We should ignore the cell we are checking

            If (y1 <> y) Then

                ' If any cell in the column has the same value as the one we are checking then
                ' mark it as invalid

                If (c(x, y1).value = value) Then
                    c(x, y1).valid = False
                End If
            End If
        Next

        Return True

    End Function

    Function check_3by3(ByVal x As Integer, ByVal y As Integer) As Boolean

        Dim x_pos As Integer
        Dim y_pos As Integer
        Dim x1 As Integer
        Dim y1 As Integer
        Dim value As Integer

        ' Get the x & y start position of the 3x3 group

        x_pos = Int(x / 3)
        x_pos = x_pos * 3
        y_pos = Int(y / 3)
        y_pos = y_pos * 3

        ' Get the value of the cell we are checking

        value = c(x, y).value

        ' Check each cell in the current 3x3 group

        For y1 = y_pos To y_pos + 2
            For x1 = x_pos To x_pos + 2

                ' We should ignore the cell we are checking

                If (x1 <> x And y1 <> y) Then

                    ' If any cell in the 3x3 group has the same value as the one we are checking then
                    ' mark it as invalid

                    If (c(x1, y1).value = value) Then
                        c(x1, y1).valid = False
                    End If

                End If
            Next
        Next

        Return True

    End Function

    Function get_number() As Integer

        Dim data As String
        Dim number As Integer
        Dim valid As Boolean

        valid = False

        Do Until valid = True

            data = InputBox("Enter number 1-9 or leave blank to reset cell: ", "Number?")
            If (Trim(data) = "") Then
                number = 0
                valid = True
            Else
                ' Checking that the data entered is numeric
                If IsNumeric(data) Then
                    number = Int(data)
                    ' If the data is numeric then it checks to see if it is BETWEEN 0-10 
                    ' if not outputs an error message 
                    If (number <= 0 Or number >= 10) Then
                        MsgBox("Enter a number between 1 and 9!", MsgBoxStyle.Exclamation, "Error")
                    Else
                        valid = True
                    End If
                Else
                    ' If the user inputs data that is not numeric an error is shown 
                    MsgBox("Enter a number!", MsgBoxStyle.Exclamation, "Error")
                End If

            End If

        Loop

        Return number

    End Function

    Function create_game()
        Dim difficulty As Integer

        ' Setting the difficulty of the Sudoku board
        difficulty = get_difficulty()

        ' Set the values for the cells that are part of the game
        Select Case difficulty
            Case Is = 1

                ' Gives the form a title
                Me.Text = "Sudoku Challenge - Easy"

                set_cell(0, 1, 6)
                set_cell(0, 2, 9)
                set_cell(0, 6, 4)
                set_cell(0, 7, 1)
                set_cell(1, 0, 5)
                set_cell(1, 1, 4)
                set_cell(1, 3, 7)
                set_cell(1, 4, 8)
                set_cell(1, 5, 1)
                set_cell(1, 7, 6)
                set_cell(1, 8, 9)
                set_cell(2, 3, 9)
                set_cell(2, 5, 4)
                set_cell(3, 1, 9)
                set_cell(3, 2, 3)
                set_cell(3, 6, 8)
                set_cell(3, 7, 7)
                set_cell(4, 2, 4)
                set_cell(4, 4, 2)
                set_cell(4, 6, 9)
                set_cell(5, 1, 8)
                set_cell(5, 2, 5)
                set_cell(5, 6, 1)
                set_cell(5, 7, 4)
                set_cell(6, 3, 3)
                set_cell(6, 5, 6)
                set_cell(7, 0, 8)
                set_cell(7, 1, 2)
                set_cell(7, 3, 4)
                set_cell(7, 4, 1)
                set_cell(7, 5, 9)
                set_cell(7, 7, 3)
                set_cell(7, 8, 5)
                set_cell(8, 1, 3)
                set_cell(8, 2, 6)
                set_cell(8, 6, 7)
                set_cell(8, 7, 9)

            Case Is = 2

                ' Gives the form a title
                Me.Text = "Sudoku Challenge - Medium"

                set_cell(0, 0, 9)
                set_cell(0, 1, 5)
                set_cell(0, 6, 3)
                set_cell(0, 7, 6)
                set_cell(1, 4, 8)
                set_cell(1, 5, 3)
                set_cell(1, 6, 5)
                set_cell(1, 7, 4)
                set_cell(2, 3, 2)
                set_cell(3, 4, 7)
                set_cell(3, 5, 2)
                set_cell(3, 7, 5)
                set_cell(4, 0, 5)
                set_cell(4, 4, 9)
                set_cell(4, 8, 1)
                set_cell(5, 1, 2)
                set_cell(5, 3, 4)
                set_cell(5, 4, 6)
                set_cell(6, 5, 7)
                set_cell(7, 1, 7)
                set_cell(7, 2, 3)
                set_cell(7, 3, 5)
                set_cell(7, 4, 2)
                set_cell(8, 1, 4)
                set_cell(8, 2, 5)
                set_cell(8, 7, 1)
                set_cell(8, 8, 8)

            Case Is = 3

                ' Gives the form a title
                Me.Text = "Sudoku Challenge - Hard"

                set_cell(0, 8, 2)
                set_cell(1, 0, 4)
                set_cell(1, 2, 5)
                set_cell(1, 4, 2)
                set_cell(2, 2, 7)
                set_cell(2, 5, 9)
                set_cell(3, 0, 2)
                set_cell(3, 5, 6)
                set_cell(3, 6, 3)
                set_cell(3, 7, 1)
                set_cell(4, 0, 3)
                set_cell(4, 5, 2)
                set_cell(4, 6, 4)
                set_cell(4, 7, 9)
                set_cell(5, 5, 5)
                set_cell(6, 4, 8)
                set_cell(7, 2, 6)
                set_cell(7, 6, 7)
                set_cell(8, 3, 1)
                set_cell(8, 4, 9)
                set_cell(8, 6, 2)
                set_cell(8, 7, 3)
        End Select
        
        Return True

    End Function

    Function set_cell(ByVal y As Integer, ByVal x As Integer, ByVal value As Integer)

        ' Set the cell value

        c(x, y).value = value

        ' Set the button to disabled so that it isn't clickable

        c(x, y).button.Enabled = False

        ' Set the button text to the value

        c(x, y).button.Text = value

        ' Set the button background colour to light green

        c(x, y).button.BackColor = Color.LightGreen

        Return True

    End Function

    Function get_difficulty() As Integer

        Dim data As String
        Dim number As Integer
        Dim valid As Boolean
        Dim message As Object

        valid = False

        Do Until valid = True
            ' Asking the user to select the difficulty that they wish
            data = InputBox("Enter number 1: Easy, 2: Medium, 3: Hard: ", "Difficulty?")
            If (Trim(data) = "") Then
                number = 0
                valid = True
            Else

                ' Checking that the data entered is numeric
                If IsNumeric(data) Then
                    number = Int(data)
                    ' If the data is numeric then it checks to see if it is BETWEEN 0-4 
                    ' if not outputs an error message 
                        If (number <= 0 Or number >= 4) Then
                        message = MsgBox("Enter a number between 1 and 3!", MsgBoxStyle.Exclamation, "Error")
                        Else
                        valid = True

                        End If
                    Else
                        ' If the user inputs data that is not numeric an error is shown 
                        MsgBox("Enter a number!", MsgBoxStyle.Exclamation, "Error")
                    End If

                End If

        Loop


        Return number
    End Function

End Class
