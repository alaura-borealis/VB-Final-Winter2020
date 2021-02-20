' Laura Waite CS 115 Winter 2020

Option Strict On
Option Explicit On

' Import statement, required for Form Closing event
Imports System.ComponentModel

Public Class frmGameBoard
    '-------------------------------------------------------------------------- 
    ' Environmental Board Game - Game Board Form
    ' Laura Waite
    ' CS 115 Winter 2020 
    ' Date: 2020-03-20 
    ' Definition: A board game for two players - tokens move around the board toward finish line
    ' Tokens are moved back/forward based on answer for multiple choice quiz questions
    '--------------------------------------------------------------------------

    ' Parallel arrays for game board squares
    Dim lblArray(99) As Label           ' Array of labels/game squares on board
    Dim intSquare(99) As Integer        ' Array for the number associated to game squares, corresponds to lblArray entries
    Dim blnOccupied(99) As Boolean      ' Contains True if square occupied - False if not, corresponds to lblArray entries

    ' Matching elements from lblArray for Roots special squares
    Dim intRootsSpecial() As Integer = {1, 2, 4, 6, 7, 11, 17, 21, 27}

    ' Matching elements from lblArray for Roots squares
    Dim intRoots() As Integer = {0, 3, 5, 8, 9, 10, 12, 13, 14,
                                 15, 16, 18, 19, 20, 22, 23, 24,
                                 25, 26, 28, 29}

    ' Matching elements from lblArray for Branches special squares
    Dim intBranchesSpecial() As Integer = {41, 43, 51, 53, 61, 65, 68, 71, 72, 76, 78}

    ' Matching elements from lblArray for Branches squares
    Dim intBranches() As Integer = {30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
                                    42, 44, 45, 46, 47, 48, 49, 50, 52, 54, 55,
                                    56, 57, 58, 59, 60, 62, 63, 64, 66, 67, 69,
                                    70, 73, 74, 75, 77, 79}

    ' Matching elements from lblArray for Treetop special squares
    Dim intTreetopSpecial() As Integer = {82, 85, 86, 95, 96}

    ' Matching elements from lblArray for Treetop squares - 97, 98, & 99 purposefully omitted
    Dim intTreetop() As Integer = {80, 81, 83, 84, 87, 88, 89, 90, 91, 92, 93, 94}

    ' Integers for die rolls
    Dim intFirstRoll As Integer
    Dim intRoll As Integer

    ' Counter for number of turns taken
    Dim intP1Turn As Integer = 0
    Dim intP2Turn As Integer = 0

    ' Player location
    Dim intP1Square As Integer = 0           ' Location for Player 1 - cannot be null
    Dim intP2Square As Integer = 0           ' Location for Player 2 - cannot be null
    Dim intP1NewSquare As Integer = vbNull   ' New location for Player 1 - null if not used
    Dim intP2NewSquare As Integer = vbNull   ' New location for Player 2 - null if not used

    ' Used for placing 2 players in same square, True if both players occupy same square
    Dim blnDoubleOccupancy As Boolean = False

    ' Variables for winner
    Dim strWinner As String = String.Empty
    Dim blnIsWinner As Boolean

    ' Boolean for playing again/resetting the game
    Dim blnPlayAgain As Boolean = False

    ' Multiple choice quiz form
    Dim frmMultipleChoice As New frmMultipleChoice

    Private Sub frmGameBoard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '-------------------------------------------------------------------------- 
        ' Subroutine - frmGameBoard_Load
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: On-load event handler for Game Board
        '--------------------------------------------------------------------------
        ' Define square size as 80 x 80 pixels
        Const intSQUARE_SIZE As Integer = 80

        ' Define font for labels
        Dim myfont As New Font("Sans Serif", 26, FontStyle.Regular)

        For intCount = 0 To 99
            ' Define label size, name, visibility, color, text alignment, and font
            lblArray(intCount) = New Label
            lblArray(intCount).Size = New System.Drawing.Size(intSQUARE_SIZE, intSQUARE_SIZE)
            lblArray(intCount).Name = "lbl" & intCount
            lblArray(intCount).Visible = True
            lblArray(intCount).BackColor = System.Drawing.Color.Transparent
            lblArray(intCount).ForeColor = System.Drawing.Color.Magenta
            lblArray(intCount).TextAlign = ContentAlignment.MiddleCenter
            lblArray(intCount).Font = myfont

            ' Set location of labels, tokens start in lower left and switch directions each row to top right
            Select Case intCount
                Case Is < 10
                    lblArray(intCount).Location = New System.Drawing.Point(intSQUARE_SIZE * intCount, 720)
                Case 10 To 19
                    lblArray(intCount).Location = New System.Drawing.Point(720 - intSQUARE_SIZE * (intCount - 10), 640)
                Case 20 To 29
                    lblArray(intCount).Location = New System.Drawing.Point(intSQUARE_SIZE * (intCount - 20), 560)
                Case 30 To 39
                    lblArray(intCount).Location = New System.Drawing.Point(720 - intSQUARE_SIZE * (intCount - 30), 480)
                Case 40 To 49
                    lblArray(intCount).Location = New System.Drawing.Point(intSQUARE_SIZE * (intCount - 40), 400)
                Case 50 To 59
                    lblArray(intCount).Location = New System.Drawing.Point(720 - intSQUARE_SIZE * (intCount - 50), 320)
                Case 60 To 69
                    lblArray(intCount).Location = New System.Drawing.Point(intSQUARE_SIZE * (intCount - 60), 240)
                Case 70 To 79
                    lblArray(intCount).Location = New System.Drawing.Point(720 - intSQUARE_SIZE * (intCount - 70), 160)
                Case 80 To 89
                    lblArray(intCount).Location = New System.Drawing.Point(intSQUARE_SIZE * (intCount - 80), 80)
                Case 90 To 99
                    lblArray(intCount).Location = New System.Drawing.Point(720 - intSQUARE_SIZE * (intCount - 90), 0)
            End Select

            ' Add labels to form
            Controls.Add(lblArray(intCount))
        Next

        'Initialize blnOccupied
        For intX = 0 To blnOccupied.Length - 1
            blnOccupied(intX) = False
        Next

        ' Randomize die roll
        Randomize()

        ' P1 goes first, enable and focus button
        blnP1Turn = True
        btnP1.Enabled = True
        btnP1.Focus()

        ' P2 not ready, disable button
        blnP2Turn = False
        btnP2.Enabled = False

        ' Finish loading form
        Me.Visible = True

        ' Show the rules
        MsgBox("Players roll a die to travel up the tree. If you land on a special square, you'll get a chance to answer a quiz question. If you're right, you might climb further up the tree. If you're wrong, watch out-- you might fall! The player who reaches the top first wins. Ready to play?",, "The Rules")

    End Sub

    Private Sub btnP1_Click(sender As Object, e As EventArgs) Handles btnP1.Click
        '-------------------------------------------------------------------------- 
        ' Subroutine - btnP1_Click
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Click event handler for Player 1 button
        '--------------------------------------------------------------------------
        ' Game play continues until Player 1 or Player 2 reaches finish

        ' Add one to turns taken counter
        intP1Turn += 1

        If intP1Turn = 1 Then
            ' Player's first turn

            ' Integer for first die roll - random number 0-5 (6 indexes)
            intFirstRoll = CInt(Int(Rnd() * 6))

            ' Convert first roll
            intRoll = intFirstRoll

        Else
            ' Subsequent turns

            ' Integer for die roll - random number 1-6
            intRoll = CInt(Int(Rnd() * 6 + 1))

        End If

        ' Ensure index is in bounds
        If (intP1Square + intRoll) < 99 Then

            If intP1Turn = 1 Then
                ' Show first die roll
                MsgBox("You rolled a " & (intFirstRoll + 1).ToString & "!",, "")

            Else
                ' Show die roll
                MsgBox("You rolled a " & intRoll.ToString & "!",, "")

                ' Remove token from previous location
                If blnDoubleOccupancy Then

                    ' Double occupied labels reset to only P2 token
                    lblArray(intP1Square).Text = ChrW(9733)
                    blnDoubleOccupancy = False
                Else
                    ' Single occupied square cleared 
                    lblArray(intP1Square).Text = String.Empty
                    blnOccupied(intP1Square) = False

                End If
            End If

            ' Index is in bounds, add die roll to current location
            intP1Square += intRoll

            ' Check for double occupancy
            If intP1Square <> intP2Square Then

                ' Square not already occupied
                blnDoubleOccupancy = False

                ' Output P1 token to new location
                lblArray(intP1Square).Text = ChrW(9734)

                ' Set as occupied
                blnOccupied(intP1Square) = True

            ElseIf intP1Square = intP2Square And intP1Turn = 1 Then

                ' Center P1 token on first turn
                blnDoubleOccupancy = False

                ' Output P1 token to new location
                lblArray(intP1Square).Text = ChrW(9734)

                ' Set as occupied
                blnOccupied(intP1Square) = True

            Else
                ' Square already occupied
                blnDoubleOccupancy = True

                ' Output P1 token on new line in label
                lblArray(intP1Square).Text &= ControlChars.CrLf & ChrW(9734)

            End If

            ' Check if location is in roots special square
            For Each intX As Integer In intRootsSpecial

                ' If location is roots special square
                If intX = intP1Square Then

                    ' Pass location to Roots function, return new location
                    intP1NewSquare = Roots(intP1Square)

                    ' Remove token from previous location
                    If blnDoubleOccupancy Then

                        ' Double occupied labels reset to only P2 token
                        lblArray(intP1Square).Text = ChrW(9733)
                        blnDoubleOccupancy = False
                    Else
                        ' Single occupied square cleared 
                        lblArray(intP1Square).Text = String.Empty
                        blnOccupied(intP1Square) = False

                    End If

                    ' Check for double occupancy
                    If intP1NewSquare <> intP2Square Then

                        ' Square not already occupied
                        blnDoubleOccupancy = False

                        ' Output P1 token to new location
                        lblArray(intP1NewSquare).Text = ChrW(9734)

                        ' Set blnOccupied to True
                        blnOccupied(intP1NewSquare) = True

                    Else
                        ' Square already occupied
                        blnDoubleOccupancy = True

                        ' Output P1 token on new line in label
                        lblArray(intP1NewSquare).Text &= ControlChars.CrLf & ChrW(9734)

                    End If

                End If
            Next

            ' Check if location is in branches special square
            For Each intX As Integer In intBranchesSpecial

                ' If location is branches special square
                If intX = intP1Square Then

                    ' Pass location to Branches function, return new location
                    intP1NewSquare = Branches(intP1Square)

                    ' Remove token from previous location
                    If blnDoubleOccupancy Then

                        ' Double occupied labels reset to only P2 token
                        lblArray(intP1Square).Text = ChrW(9733)
                        blnDoubleOccupancy = False
                    Else
                        ' Single occupied square cleared 
                        lblArray(intP1Square).Text = String.Empty
                        blnOccupied(intP1Square) = False

                    End If

                    ' Check for double occupancy
                    If intP1NewSquare <> intP2Square Then

                        ' Square not already occupied
                        blnDoubleOccupancy = False

                        ' Output P1 token to new location
                        lblArray(intP1NewSquare).Text = ChrW(9734)

                        ' Set blnOccupied to True
                        blnOccupied(intP1NewSquare) = True

                    Else
                        ' Square already occupied
                        blnDoubleOccupancy = True

                        ' Output P1 token on new line in label
                        lblArray(intP1NewSquare).Text &= ControlChars.CrLf & ChrW(9734)

                    End If

                End If
            Next

            ' Check if location is in treetop special square
            For Each intX As Integer In intTreetopSpecial

                ' If location is treetop special square
                If intX = intP1Square Then

                    ' Pass to Treetop function, return new location
                    intP1NewSquare = TreeTop(intP1Square)

                    ' Remove token from previous location
                    If blnDoubleOccupancy Then

                        ' Double occupied labels reset to only P2 token
                        lblArray(intP1Square).Text = ChrW(9733)
                        blnDoubleOccupancy = False
                    Else
                        ' Single occupied square cleared 
                        lblArray(intP1Square).Text = String.Empty
                        blnOccupied(intP1Square) = False

                    End If

                    ' Check for double occupancy
                    If intP1NewSquare <> intP2Square Then

                        ' Square not already occupied
                        blnDoubleOccupancy = False

                        ' Output P1 token to new location
                        lblArray(intP1NewSquare).Text = ChrW(9734)

                        ' Set blnOccupied to True
                        blnOccupied(intP1NewSquare) = True

                    Else
                        ' Square already occupied
                        blnDoubleOccupancy = True

                        ' Output P1 token on new line in label
                        lblArray(intP1NewSquare).Text &= ControlChars.CrLf & ChrW(9734)

                    End If

                End If
            Next

            ' Once token reaches final location for turn
            If intP1NewSquare <> vbNull Then
                ' Discard "new location" variable
                intP1Square = intP1NewSquare
                intP1NewSquare = vbNull
                blnOccupied(intP1Square) = True

            End If

            ' Ensure non-occupied labels are cleared
            For intX = 0 To blnOccupied.Length - 1
                If blnOccupied(intX) = False And blnDoubleOccupancy = False Then

                    lblArray(intX).Text = String.Empty

                End If
            Next

            ' Change turns
            Call ChangeTurns()

        ElseIf (intP1Square + intRoll) = 99 Then

            ' Show die roll
            MsgBox("You rolled a " & intRoll.ToString & "!",, "")

            ' Remove token from previous location
            If blnDoubleOccupancy Then

                ' Double occupied labels reset to only P2 token
                lblArray(intP1Square).Text = ChrW(9733)
                blnDoubleOccupancy = False
            Else
                ' Single occupied square cleared 
                lblArray(intP1Square).Text = String.Empty
                blnOccupied(intP1Square) = False

            End If

            ' Add die roll to current location
            intP1Square += intRoll

            ' Output P1 token to finish line
            lblArray(intP1Square).Text = ChrW(9734)

            ' Player 1 reached finish, call GameOver
            Call GameOver("Player 1")

        Else
            ' Index out of bounds
            intP1Square = intP1Square

            ' Show die roll and roll needed to reach finish
            MsgBox("You rolled a " & intRoll.ToString & "!" &
                       ControlChars.CrLf & "You need to roll a " &
                       CStr(99 - intP1Square) & " to win.",, "")

            ' Check for double occupancy
            If intP1Square <> intP2Square Then

                ' Square not already occupied
                blnDoubleOccupancy = False

                ' Output P1 token to same location
                lblArray(intP1Square).Text = ChrW(9734)

                ' Set as occupied
                blnOccupied(intP1Square) = True

            Else
                ' Square already occupied
                blnDoubleOccupancy = True

                ' Output P1 token on new line in label
                lblArray(intP1Square).Text &= ControlChars.CrLf & ChrW(9734)

            End If

            ' Change turns
            Call ChangeTurns()

        End If

    End Sub

    Private Sub btnP2_Click(sender As Object, e As EventArgs) Handles btnP2.Click
        '-------------------------------------------------------------------------- 
        ' Subroutine - btnP2_Click
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Click event handler for Player 2 button
        '--------------------------------------------------------------------------

        ' Game play continues until Player 1 or Player 2 reaches finish

        ' Add one to turns taken counter
        intP2Turn += 1

        If intP2Turn = 1 Then
            ' Player's first turn

            ' Integer for first die roll - random number 0-5 (6 indexes)
            intFirstRoll = CInt(Int(Rnd() * 6))

            ' Convert first roll
            intRoll = intFirstRoll

        Else
            ' Subsequent turns

            ' Integer for die roll - random number 1-6
            intRoll = CInt(Int(Rnd() * 6 + 1))

        End If

        ' After roll has been added, ensure index is in bounds
        If (intP2Square + intRoll) < 99 Then

            If intP2Turn = 1 Then
                ' Show first die roll
                MsgBox("You rolled a " & (intFirstRoll + 1).ToString & "!",, "")

            Else
                ' Show die roll
                MsgBox("You rolled a " & intRoll.ToString & "!",, "")

                ' Remove token from previous location
                If blnDoubleOccupancy Then

                    ' Double occupied labels reset to only P1 token
                    lblArray(intP2Square).Text = ChrW(9734)
                    blnDoubleOccupancy = False
                Else
                    ' Single occupied square cleared 
                    lblArray(intP2Square).Text = String.Empty
                    blnOccupied(intP2Square) = False

                End If
            End If

            ' Index is in bounds, add die roll to current location
            intP2Square += intRoll

            ' Check for double occupancy
            If intP2Square <> intP1Square Then

                ' Square not already occupied
                blnDoubleOccupancy = False

                ' Output P2 token to new location
                lblArray(intP2Square).Text = ChrW(9733)

                ' Set as occupied
                blnOccupied(intP2Square) = True

            Else
                ' Square already occupied
                blnDoubleOccupancy = True

                ' Output P2 token on new line in label
                lblArray(intP2Square).Text &= ControlChars.CrLf & ChrW(9733)

            End If

            ' Check if location is in roots special square
            For Each intX As Integer In intRootsSpecial

                ' If location is roots special square
                If intX = intP2Square Then

                    ' Pass location to Roots function, return new location
                    intP2NewSquare = Roots(intP2Square)

                    ' Remove token from previous location
                    If blnDoubleOccupancy Then

                        ' Double occupied labels reset to only P1 token
                        lblArray(intP2Square).Text = ChrW(9734)
                        blnDoubleOccupancy = False
                    Else
                        ' Single occupied square cleared 
                        lblArray(intP2Square).Text = String.Empty
                        blnOccupied(intP2Square) = False

                    End If

                    ' Check for double occupancy
                    If intP2NewSquare <> intP1Square Then

                        ' Square not already occupied
                        blnDoubleOccupancy = False

                        ' Output P2 token to new location
                        lblArray(intP2NewSquare).Text = ChrW(9733)

                        ' Set blnOccupied to True
                        blnOccupied(intP2NewSquare) = True

                    Else
                        ' Square already occupied
                        blnDoubleOccupancy = True

                        ' Output P2 token on new line in label
                        lblArray(intP2NewSquare).Text &= ControlChars.CrLf & ChrW(9733)

                    End If

                End If
            Next

            ' Check if location is in branches special square
            For Each intX As Integer In intBranchesSpecial

                ' If location is branches special square
                If intX = intP2Square Then

                    ' Pass location to Branches function, return new location
                    intP2NewSquare = Branches(intP2Square)

                    ' Remove token from previous location
                    If blnDoubleOccupancy Then

                        ' Double occupied labels reset to only P1 token
                        lblArray(intP2Square).Text = ChrW(9734)
                        blnDoubleOccupancy = False
                    Else
                        ' Single occupied square cleared 
                        lblArray(intP2Square).Text = String.Empty
                        blnOccupied(intP2Square) = False

                    End If

                    ' Check for double occupancy
                    If intP2NewSquare <> intP1Square Then

                        ' Square not already occupied
                        blnDoubleOccupancy = False

                        ' Output P2 token to new location
                        lblArray(intP2NewSquare).Text = ChrW(9733)

                        ' Set blnOccupied to True
                        blnOccupied(intP2NewSquare) = True

                    Else
                        ' Square already occupied
                        blnDoubleOccupancy = True

                        ' Output P2 token on new line in label
                        lblArray(intP2NewSquare).Text &= ControlChars.CrLf & ChrW(9733)

                    End If

                End If
            Next

            ' Check if location is in treetop special square
            For Each intX As Integer In intTreetopSpecial

                ' If location is treetop special square
                If intX = intP2Square Then

                    ' Pass to Treetop function, return new location
                    intP2NewSquare = TreeTop(intP2Square)

                    ' Remove token from previous location
                    If blnDoubleOccupancy Then

                        ' Double occupied labels reset to only P1 token
                        lblArray(intP2Square).Text = ChrW(9734)
                        blnDoubleOccupancy = False
                    Else
                        ' Single occupied square cleared 
                        lblArray(intP2Square).Text = String.Empty
                        blnOccupied(intP2Square) = False

                    End If

                    ' Check for double occupancy
                    If intP2NewSquare <> intP1Square Then

                        ' Square not already occupied
                        blnDoubleOccupancy = False

                        ' Output P2 token to new location
                        lblArray(intP2NewSquare).Text = ChrW(9733)

                        ' Set blnOccupied to True
                        blnOccupied(intP2NewSquare) = True
                    Else
                        ' Square already occupied
                        blnDoubleOccupancy = True

                        ' Output P2 token on new line in label
                        lblArray(intP2NewSquare).Text &= ControlChars.CrLf & ChrW(9733)

                    End If

                End If
            Next

            ' Once token reaches final location for turn
            If intP2NewSquare <> vbNull Then

                ' Discard "new location" variable
                intP2Square = intP2NewSquare
                intP2NewSquare = vbNull
                blnOccupied(intP2Square) = True

            End If

            ' Ensure non-occupied labels are cleared
            For intX = 0 To blnOccupied.Length - 1
                If blnOccupied(intX) = False And blnDoubleOccupancy = False Then

                    lblArray(intX).Text = String.Empty

                End If
            Next

            ' Change turns
            Call ChangeTurns()

        ElseIf (intP2Square + intRoll) = 99 Then

            ' Show die roll
            MsgBox("You rolled a " & intRoll.ToString & "!",, "")

            ' Remove token from previous location

            If blnDoubleOccupancy Then

                ' Double occupied labels reset to only P1 token
                lblArray(intP2Square).Text = ChrW(9734)
                blnDoubleOccupancy = False
            Else
                ' Single occupied square cleared 
                lblArray(intP2Square).Text = String.Empty
                blnOccupied(intP2Square) = False

            End If

            ' Add die roll to current location
            intP2Square += intRoll

            ' Output P2 token to finish line
            lblArray(intP2Square).Text = ChrW(9733)

            ' Player 2 reached finish, call GameOver
            Call GameOver("Player 2")

        Else
            ' Index out of bounds
            intP2Square = intP2Square

            ' Show die roll and roll needed to reach finish
            MsgBox("You rolled a " & intRoll.ToString & "!" &
                   ControlChars.CrLf & "You need to roll a " &
                   CStr(99 - intP2Square) & " to win.",, "")

            ' Check for double occupancy
            If intP2Square <> intP1Square Then

                ' Square not already occupied
                blnDoubleOccupancy = False

                ' Output P2 token to same location
                lblArray(intP2Square).Text = ChrW(9733)

                ' Set as occupied
                blnOccupied(intP2Square) = True

            Else
                ' Square already occupied
                blnDoubleOccupancy = True

                ' Output P2 token on new line in label
                lblArray(intP2Square).Text &= ControlChars.CrLf & ChrW(9733)

            End If

            ' Change turns
            Call ChangeTurns()

        End If

    End Sub

    Private Function Roots(intSquare As Integer) As Integer
        '-------------------------------------------------------------------------- 
        ' Function - Roots
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Function for square behavior in intRootsSpecial array
        '--------------------------------------------------------------------------

        ' Moves player into greater root if answer is correct, does nothing if incorrect

        ' Open MultipleChoice form
        frmMultipleChoice.ShowDialog()

        ' Variable for random location in root array
        Dim intRoot As Integer = CInt(Int(Rnd() * 19))

        ' Variable to hold new location for loop
        Dim intNewSquare As Integer

        ' Set rules for correct versus incorrect answer
        If blnCorrect Then

            ' If answer is correct, verify new location is higher than original
            If intRoots(intRoot) > intSquare Then

                ' New location is higher, move token
                intSquare = intRoots(intRoot)

                ' Return token location
                Return intSquare

            ElseIf (blnCorrect = True Or blnCorrect = False) And (blnP1LoseATurn Or blnP2LoseATurn) Then

                ' If Player lost a turn on this question, no movement
                intSquare = intSquare
                Return intSquare

            Else
                ' Otherwise, add one to index until new location is greater

                intNewSquare = intRoots(intRoot)

                Do While intNewSquare < intSquare

                    intRoot += 1
                    intNewSquare = intRoots(intRoot)
                Loop

                intSquare = intNewSquare

                ' Return token location
                Return intSquare

            End If

        Else
            ' If answer is incorrect, do nothing
            intSquare = intSquare

            ' Return token location
            Return intSquare
        End If

    End Function

    Private Function Branches(intSquare As Integer) As Integer
        '-------------------------------------------------------------------------- 
        ' Function - Branches
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Function for square behavior in intBranchesSpecial array
        '--------------------------------------------------------------------------

        ' Moves player into greater branch if answer is correct, moves to lesser branch if incorrect

        ' Open MultipleChoice form
        frmMultipleChoice.ShowDialog()

        ' Variable for random location in branches array
        Dim intBranch As Integer = CInt(Int(Rnd() * 37))

        ' Variable to hold new location for loop
        Dim intNewSquare As Integer

        ' Set rules for correct versus incorrect answer
        If blnCorrect Then

            ' If answer is correct, verify new location is higher than original
            If intBranches(intBranch) > intSquare Then

                ' New location is higher, move token
                intSquare = intBranches(intBranch)

                ' Return token location
                Return intSquare

            Else
                ' Otherwise, add one to index until new location is higher

                Do While intNewSquare < intSquare

                    intBranch += 1
                    intNewSquare = intBranches(intBranch)
                Loop

                intSquare = intNewSquare

                ' Return token location
                Return intSquare

            End If
        ElseIf (blnCorrect = True Or blnCorrect = False) And (blnP1LoseATurn Or blnP2LoseATurn) Then

            ' If Player lost a turn on this question, no movement
            intSquare = intSquare
            Return intSquare

        Else
            ' If answer is incorrect, verify new location is lower than original

            If intBranches(intBranch) < intSquare Then

                ' New location is lower, move token
                intSquare = intBranches(intBranch)

                ' Return token location
                Return intSquare

            Else
                ' Otherwise, subtract one from index until new location is lower
                intNewSquare = intBranches(intBranch)

                Do While intNewSquare > intSquare

                    intBranch -= 1
                    intNewSquare = intBranches(intBranch)
                Loop

                intSquare = intNewSquare

                ' Return token location
                Return intSquare


            End If
        End If

    End Function

    Private Function TreeTop(intSquare As Integer) As Integer
        '-------------------------------------------------------------------------- 
        ' Function - Treetop
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Function for square behavior in intTreetopSpecial array
        '--------------------------------------------------------------------------
        ' Moves player into branches if answer is correct, does nothing if incorrect

        ' Open MultipleChoice form
        frmMultipleChoice.ShowDialog()

        ' Variable for random location in branches array
        Dim intTop As Integer = CInt(Int(Rnd() * 11))

        ' Variable to hold new location for loop
        Dim intNewSquare As Integer

        ' Set rules for correct versus incorrect answer
        If blnCorrect Then

            ' If answer is correct, do nothing
            intSquare = intSquare

            ' Return token location
            Return intSquare

        ElseIf (blnCorrect = True Or blnCorrect = False) And (blnP1LoseATurn Or blnP2LoseATurn) Then

            ' If Player lost a turn on this question, no movement
            intSquare = intSquare
            Return intSquare

        Else
            ' If answer is incorrect, verify new location is lower than original
            If intTreetop(intTop) < intSquare Then

                ' New location is lower, move token to new location
                intSquare = intTreetop(intTop)

                ' Return token location
                Return intSquare

            Else
                ' Otherwise, subtract one from index until new location is lower

                intNewSquare = intTreetop(intTop)

                Do While intNewSquare > intSquare

                    intTop -= 1
                    intNewSquare = intTreetop(intTop)
                Loop

                intSquare = intNewSquare

                ' Return token location
                Return intSquare

            End If
        End If

    End Function

    Sub ChangeTurns()
        '-------------------------------------------------------------------------- 
        ' Subroutine - ChangeTurns
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Subroutine for changing turns between players
        '--------------------------------------------------------------------------
        ' Define turn change rules
        If blnP1Turn And Not blnP2LoseATurn Then

            ' Player 1 has completed their turn, Player 2 did not lose turn at last turn
            blnP1Turn = False

            ' Disable Player 1 button
            btnP1.Enabled = False

            ' Set Player 2 toggle
            blnP2Turn = True

            ' Enable Player 2 button
            btnP2.Enabled = True

            ' Set Player 2 button focus
            btnP2.Focus()

        ElseIf blnP1Turn And blnP2LoseATurn Then

            ' Player 1 has completed their turn, Player 2 lost turn at last turn
            blnP2Turn = False

            ' Player 2 button remains disabled
            btnP2.Enabled = False

            ' Player 1 turn toggle stays set
            blnP1Turn = True

            ' Player 1 button stays enabled
            btnP1.Enabled = True

            ' Set Player 1 button focus
            btnP1.Focus()

            ' Reset after lost turn
            blnP2LoseATurn = False

        ElseIf blnP2Turn And blnP1LoseATurn Then

            ' Player 2 has completed their turn, Player 1 lost turn at last turn
            blnP1Turn = False

            ' Player 1 button remains disabled
            btnP1.Enabled = False

            ' Player 2 toggle stays set
            blnP2Turn = True

            ' Player 2 button stays enabled
            btnP2.Enabled = True

            ' Set Player 2 button focus
            btnP2.Focus()

            ' Reset after lost turn
            blnP1LoseATurn = False

        Else
            ' Player 2 has completed their turn, Player 1 did not lose turn at last turn
            blnP2Turn = False

            ' Disable Player 2 button
            btnP2.Enabled = False

            ' Set Player 1 toggle
            blnP1Turn = True

            ' Enable Player 1 button
            btnP1.Enabled = True

            ' Set Player 1 button focus
            btnP1.Focus()

        End If
    End Sub

    Sub GameOver(strWinner As String)
        '----------------------------------------------------------------------------
        ' Subroutine - GameOver
        ' Laura Waite
        ' CS 115 Winter 2020
        ' Date: 2020-03-20
        ' Description: Subroutine to declare winner and offer replay
        '-----------------------------------------------------------------------------
        ' Variable to hold user response to message box question
        Dim enumResponse As MsgBoxResult

        ' Declare winner, offer choice to play again
        blnIsWinner = True
        enumResponse = MsgBox((strWinner & " wins!" &
                           ControlChars.CrLf & "Do you want to play again?"), MsgBoxStyle.YesNo, "Hooray!")

        ' Set play again boolean based on user selection
        If enumResponse = MsgBoxResult.Yes Then

            ' User clicked Yes
            blnPlayAgain = True

        ElseIf enumResponse = MsgBoxResult.No Then

            ' User clicked No
            blnPlayAgain = False

        End If

        ' Call PlayAgain function
        PlayAgain(blnPlayAgain)

    End Sub

    Sub PlayAgain(blnReplay As Boolean)
        '-------------------------------------------------------------------------- 
        ' Function - PlayAgain
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Function that runs when a player reaches finish
        '--------------------------------------------------------------------------
        ' Evaluate argument passed by GameOver subroutine
        If blnReplay Then

            'User selected "Yes", reinitialize blnOccupied and clear labels
            For intX = 0 To 99
                blnOccupied(intX) = False
                lblArray(intX).Text = String.Empty
            Next

            ' Reset player turn counters
            intP1Turn = 0
            intP2Turn = 0

            ' Reset player locations
            intP1Square = 0
            intP2Square = 0
            intP1NewSquare = vbNull
            intP2NewSquare = vbNull

            ' Reset winner variables
            blnIsWinner = False
            strWinner = ""

            ' P1 goes first, enable and focus button
            blnP1Turn = True
            btnP1.Enabled = True
            btnP1.Focus()

            ' P2 not ready, disable button
            blnP2Turn = False
            btnP2.Enabled = False

        Else
            ' User selected "No" - close application
            Me.Close()

        End If
    End Sub

    Private Sub frmGameBoard_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        '-------------------------------------------------------------------------- 
        ' Subroutine - frmGameBoard_Closing
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: On close event handler for Environmental Board Game application
        '--------------------------------------------------------------------------
        ' Prevent redundant GameOver messaging
        If blnIsWinner = False Then

            ' User selected "Close" button or Alt-F4 - game ended before winner declared - offer new gam
            Dim enumReset As MsgBoxResult = MsgBox("Would you like to start a new game instead?", MsgBoxStyle.YesNo, "Leaving so soon?")

            If enumReset = MsgBoxResult.Yes Then

                ' User has accepted replay
                blnPlayAgain = True

                ' Cancel form closing event
                e.Cancel = True

                'Call PlayAgain to reset gameplay
                PlayAgain(blnPlayAgain)

            Else
                ' User has declined replay
                Me.Visible = False
                MsgBox("Thanks for playing!",, "Bye!")

            End If

        Else
            ' Winner was declared and user has declined replay
            Me.Visible = False
            MsgBox("Thanks for playing!",, "Bye!")

        End If
    End Sub

End Class