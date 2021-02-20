' Laura Waite CS 115 Winter 2020

Option Strict On
Option Explicit On

Module MultiChoiceModule
    '-------------------------------------------------------------------------- 
    ' Environmental Board Game - Multi Choice Module
    ' Laura Waite
    ' CS 115 Winter 2020 
    ' Date: 2020-03-20 
    ' Definition: Module for global variables shared between application forms
    '--------------------------------------------------------------------------
    ' Declare global variables for forms

    ' Toggle between players, True for active player's turn
    Public blnP1Turn As Boolean
    Public blnP2Turn As Boolean

    ' Booleans for correct answer
    Public blnCorrect, blnAnswer As Boolean

    ' Boolean for opponent losing a turn
    Public blnP1LoseATurn As Boolean = False
    Public blnP2LoseATurn As Boolean = False

    ' Counter for number of questions asked
    Public intCounter As Integer = 0

    Public Sub QuestionResult()
        '-------------------------------------------------------------------------- 
        ' Subroutine - QuestionResult
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Public subroutine that sets the question answer to correct
        '--------------------------------------------------------------------------
        If blnAnswer = True Then
            blnCorrect = True
        Else
            blnCorrect = False
        End If
    End Sub

    Public Sub QuestionCounter()
        '-------------------------------------------------------------------------- 
        ' Subroutine - QuestionCounter
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Public subroutine to count number of questions
        ' as determined by on-load events for frmMultipleChoice
        ' Does not reset with PlayAgain subroutine
        '--------------------------------------------------------------------------
        ' Add one to question counter
        intCounter += 1
    End Sub
End Module