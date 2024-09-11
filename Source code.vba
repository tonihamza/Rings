' ======================================================
' WARNING:
' This code is split into multiple modules. Ensure that:
' 1. The code for generating random colors is placed in one module.
' 2. The code for changing cell colors is placed in a separate module.
' 3. The code for creating random colored rings is placed in another module.
' ======================================================

' ===================
' Module 1: Random Color Generation
' ===================

' Public variable to store a list of colors
Public colors(1 To 50) As Long

' ============================
' Subroutine: GenerateRandomColors
' Generates a list of random colors and stores them in the global array.
' ============================

Sub GenerateRandomColors()
    Dim i As Long
    ' Initialize the random number generator
    Randomize
    
    ' Generate random colors and store them in the global array
    For i = 1 To 50
        colors(i) = RGB(Int(256 * Rnd), Int(256 * Rnd), Int(256 * Rnd))
    Next i
    MsgBox "Done" ' Notify the user that color generation is complete
    MsgBox colors(10) ' Display a sample color for verification
End Sub

' ===================
' Module 2: Random Color Change
' ===================

' ============================
' Subroutine: ChangeCellColorRandomly
' Changes the background color of the selected cell randomly.
' ============================

Sub ChangeCellColorRandomly()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim r As Integer, g As Integer, b As Integer
    Dim startTime As Double
    Dim i As Long
    Dim numColors As Long

    ' Set the worksheet and the selected cell
    Set ws = ActiveSheet
    Set selectedCell = Selection

    ' Number of color changes
    numColors = 1000 ' Adjust the number of color changes as needed

    ' Initialize the random number generator
    Randomize

    ' Loop to change the color of the cell randomly
    For i = 1 To numColors
        ' Generate random RGB values with a step of 10
        r = Int((256 / 10) * Rnd) * 10
        g = Int((256 / 10) * Rnd) * 10
        b = Int((256 / 10) * Rnd) * 10
        
        ' Ensure values are within valid range
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255

        ' Change the color of the selected cell
        selectedCell.Interior.Color = RGB(r, g, b)
        
        ' Add a delay of 50 milliseconds
        startTime = Timer
        Do While Timer < startTime + 0.05
            DoEvents
        Loop
    Next i
End Sub

' ===================
' Module 3: Colored Rings
' ===================

' Public variable to store a list of colors
Public colors(1 To 100) As Long

' ============================
' Subroutine: GenerateRandomColors
' Generates a list of random colors and stores them in the global array.
' ============================

Sub GenerateRandomColors()
    Dim i As Long
    ' Initialize the random number generator
    Randomize
    
    ' Generate random colors and store them in the global array
    For i = 1 To 100
        colors(i) = RGB(Int(256 * Rnd), Int(256 * Rnd), Int(256 * Rnd))
    Next i
End Sub

' ============================
' Subroutine: CreateRandomColoredRings
' Creates animated expanding colored rings starting from the selected cell.
' ============================

Sub CreateRandomColoredRings()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim centerX As Long, centerY As Long
    Dim radius As Long
    Dim maxRadius As Long
    Dim x As Long, y As Long
    Dim cell As Range
    Dim generation As Long
    Dim offset As Long
    Dim startTime As Double
    
    ' Set the worksheet and the selected cell
    Set ws = ActiveSheet
    Set selectedCell = Selection
    
    ' Get the center cell coordinates
    centerX = selectedCell.Column
    centerY = selectedCell.Row
    
    ' Set the maximum radius
    maxRadius = 50  ' Adjust this value to control the maximum radius
    
    ' Generate a list of random colors
    GenerateRandomColors
    ' Ensure colors are generated
    If colors(1) = 0 Then
        MsgBox "Please run GenerateRandomColors first to initialize the color list."
        Exit Sub
    End If
    
    ' Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    
    ' Loop through each generation of rings
    For generation = 1 To 50
        offset = generation - 1
        
        ' Loop through each radius from the maximum to 1
        For radius = 1 To maxRadius Step 1
            ' Get the color for the current radius with offset
            Dim currentColor As Long
            currentColor = colors(((maxRadius - radius + offset) Mod 50) + 1)
            
            ' Loop through each cell in the current ring radius
            For x = -radius To radius
                For y = -radius To radius
                    ' Check if the cell is on the boundary of the ring
                    If Abs(x ^ 2 + y ^ 2 - radius ^ 2) <= radius Then
                        On Error Resume Next  ' Ignore errors for cells outside the worksheet
                        Set cell = ws.Cells(centerY + y, centerX + x)
                        cell.Interior.Color = currentColor ' Change the color of the cell
                        On Error GoTo 0
                    End If
                Next y
            Next x
        Next radius
        
        ' Add a delay for visualization
        startTime = Timer
        Application.ScreenUpdating = True
        Do While Timer < startTime + 0.1
            DoEvents
        Loop
        Application.ScreenUpdating = False
    Next generation
    
    ' Reset cell styles and select a specific cell
    Cells.Select
    Selection.Style = "Normal"
    Range("BC22").Select
    
    ' Turn on screen updating at the end
    Application.ScreenUpdating = True
End Sub
