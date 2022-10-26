# boardgametools
The VBA code for my Microsoft Excel Sheet to create a working timer, Dice, and coin flip features. You will need to download the Worksheet from my Wix website and assign the buttons the correct Macros for this to work.

To use this worksheet first download the Spreadsheet from my Wix website(See below). It will be under the Projects sections.
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
https://matthewloganhoran.wixsite.com/matthew-horan-s-data
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Add this Sub to your Macros:
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub Protect_worksheet()
'
' Protect_worksheet_unprotect Macro
'

'
ActiveSheet.Protect
End Sub
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Assign the following to the "Start timer" button on Sheet "Timer":
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
 Sub timer()
'

'
    ActiveSheet.Unprotect
   
     interval = Now + TimeValue("00:00:01")

     If Range("L13").Value = 0 Then Call Protect_worksheet
     If Range("L13").Value = 0 Then Exit Sub

     Range("L13") = Range("L13") - TimeValue("00:00:01")

     Application.OnTime interval, "timer"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
 End Sub
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Assign the following to the "Set 1 Minute" button on Sheet "Timer":
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub Add_1_minute()
'
' Add_1_minute Macro
'

'
   ActiveSheet.Unprotect
    Range("L13").Select
    ActiveCell.FormulaR1C1 = "12:01:00 AM"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Assign the following to the "Roll Dice" button on Sheet "Dice":
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub Roll_dice()
'
' Roll_dice Macro
'

'
    ActiveSheet.Unprotect
    Columns("W:AG").Select
    Selection.EntireColumn.Hidden = False
    Range("Z3").Select
    Selection.Copy
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.CommandBars("Office Clipboard").Visible = False
    Columns("Y:AG").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
    ActiveSheet.Protect
End Sub
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Assign the following to the "Flip Coin" button on Sheet "Coin_Flip":
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub Flip_coin()
'
' Roll_dice Macro
'

'
    ActiveSheet.Unprotect
    Range("Z3").Select
    Selection.Copy
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Application.CommandBars("Office Clipboard").Visible = False
    Range("A1").Select
    ActiveSheet.Protect
End Sub
-------------------------------------------------------------------------------------------------------------------------------------------------------------------
