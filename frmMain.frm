VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Word Puzzle"
   ClientHeight    =   7020
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAuto 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3420
      Top             =   6300
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Print Puzzle"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      ToolTipText     =   "Put the puzzle in a text file."
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      ToolTipText     =   "Thank you for using Word Puzzle."
      Top             =   1860
      Width           =   1515
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   180
      Width           =   1515
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      ToolTipText     =   "Add the random letters. "
      Top             =   1020
      Width           =   1515
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      ToolTipText     =   "Creat a new puzzle using the word in the list."
      Top             =   600
      Width           =   1515
   End
   Begin VB.Frame fraDifficulty 
      Caption         =   "Difficulty"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   3300
      TabIndex        =   13
      Top             =   4260
      Width           =   1815
      Begin VB.OptionButton optEasy 
         Caption         =   "Eas&y"
         Height          =   255
         Left            =   540
         TabIndex        =   15
         ToolTipText     =   "No diagonal words"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optHard 
         Caption         =   "&Hard"
         Height          =   195
         Left            =   540
         TabIndex        =   14
         Top             =   600
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.Frame fraRawPuzzle 
      Caption         =   "Puzzle"
      ForeColor       =   &H00FF0000&
      Height          =   5115
      Left            =   5160
      TabIndex        =   12
      Top             =   120
      Width           =   4995
      Begin RichTextLib.RichTextBox txtWord 
         Height          =   4575
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   8070
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraOfPuzzlesToCreate 
      Caption         =   "# of Puzzles to create"
      ForeColor       =   &H00FF0000&
      Height          =   1155
      Left            =   3300
      TabIndex        =   10
      Top             =   5280
      Width           =   1815
      Begin VB.CommandButton cmdAuto 
         Caption         =   "&Auto"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   22
         ToolTipText     =   "Create lots of puzzles as text files."
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtNumOfPuzzles 
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Text            =   "1"
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblNumOfPuzzles 
         Caption         =   "Welcome to Word Puzzle by John Nyhart. Currently the program is building puzzles and storing them in "
         Height          =   555
         Left            =   2100
         TabIndex        =   23
         Top             =   180
         Width           =   4635
      End
      Begin VB.Line linRed 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   3300
         X2              =   6540
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line linGrn 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         X1              =   2040
         X2              =   3240
         Y1              =   900
         Y2              =   900
      End
   End
   Begin VB.Frame fraWordCount 
      Caption         =   "Max Word Count"
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   3300
      TabIndex        =   8
      Top             =   3540
      Width           =   1815
      Begin VB.TextBox txtWordCount 
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Text            =   "100"
         ToolTipText     =   "How many words from the list for this puzzle. Words are picked by random."
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Grid"
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   3300
      TabIndex        =   3
      Top             =   2460
      Width           =   1815
      Begin VB.TextBox txtGridCol 
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Text            =   "20"
         Top             =   540
         Width           =   675
      End
      Begin VB.TextBox txtGridRows 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Text            =   "20"
         ToolTipText     =   "Min 5, Max 20"
         Top             =   180
         Width           =   675
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Cols"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   600
         Width           =   300
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Rows"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame fraInputWord 
      Caption         =   "Input Word"
      ForeColor       =   &H00FF0000&
      Height          =   6615
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3195
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove Word"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "Highlight a word and click here."
         Top             =   6180
         Width           =   1275
      End
      Begin VB.ListBox lstInputWord 
         Height          =   5325
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   780
         Width           =   2715
      End
      Begin VB.TextBox txtInputWord 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Type in word and press ENTER key."
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.Label lbl 
      Caption         =   $"frmMain.frx":03E8
      Height          =   795
      Left            =   5460
      TabIndex        =   24
      Top             =   5460
      Width           =   4575
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aryListBox() As String
Dim aryMain() As String
Dim aryTemp() As String
Dim glWordCount As Long
Dim bAutoFrameFlg As Boolean   ' false = close -> open, true = open -> close
Sub AutoExportPuzzle(sFilePath As String)
    Dim iCol As Integer
    Dim iMaxCol As Integer
    Dim iMaxRow As Integer
    Dim iRow As Integer
    Dim lCount As Long
    Dim lPointer As Long
    Dim sTmp As String

    'sFilePath = App.Path & "WordPuz.txt"
    Open sFilePath For Output As #1

    ' *** get the row and col values
    With frmMain
        iMaxRow = Val(.txtGridRows.Text)
        iMaxCol = Val(.txtGridCol.Text)
    End With '»With frmMain

    ' *** create a heading
    Print #1, "Word Puzzle by John Nyhart (www.nyhart.com)"
    Print #1, " "
    Print #1, " "

    ' *** save the finished puzzle
    For iRow = 1 To iMaxRow
        For iCol = 1 To iMaxCol
            sTmp = sTmp & aryMain(iRow, iCol) & " "
        Next '»For iCol = 1 To iMaxCol
        Print #1, sTmp
        sTmp = ""
    Next '»For iRow = 1 To iMaxRow

    Print #1, " "
    Print #1, " "
    Print #1, " "
    Print #1, "Words to find:"

    ' *** print the words to look for
    With frmMain.lstInputWord
        lCount = .ListCount - 1
        For lPointer = 0 To lCount
            Print #1, .List(lPointer)
        Next '»For lPointer = 0 To lCount
    End With '»With frmMain.lstInputWord

    Print #1, " "
    Print #1, " "
    Print #1, " "

    ' *** print the key
    For iRow = 1 To iMaxRow
        For iCol = 1 To iMaxCol
            sTmp = sTmp & aryTemp(iRow, iCol) & " "
        Next '»For iCol = 1 To iMaxCol
        Print #1, sTmp
        sTmp = ""
    Next '»For iRow = 1 To iMaxRow

    Print #1, " "
    Print #1, " "
    Print #1, "********** NOTE *****************"
    Print #1, " Import this file to Word and"
    Print #1, " change the font to Courier New."
    Print #1, "*********************************"

    Close #1
    

End Sub

Public Sub CenterForm(anyForm)

    If anyForm.WindowState = 0 Then
        anyForm.Top = (Screen.Height - anyForm.Height) / 2
        anyForm.Left = (Screen.Width - anyForm.Width) / 2
    End If

End Sub

Sub ExportPuzzle()
    Dim iCol As Integer
    Dim iMaxCol As Integer
    Dim iMaxRow As Integer
    Dim iRow As Integer
    Dim lCount As Long
    Dim lPointer As Long
    Dim sFilePath As String
    Dim sTmp As String

    sFilePath = App.Path & "WordPuz.txt"
    Open sFilePath For Output As #1

    ' *** get the row and col values
    With frmMain
        iMaxRow = Val(.txtGridRows.Text)
        iMaxCol = Val(.txtGridCol.Text)
    End With '»With frmMain

    ' *** create a heading
    Print #1, "Word Puzzle by John Nyhart (www.nyhart.com)"
    Print #1, " "
    Print #1, " "

    ' *** save the finished puzzle
    For iRow = 1 To iMaxRow
        For iCol = 1 To iMaxCol
            sTmp = sTmp & aryMain(iRow, iCol) & " "
        Next '»For iCol = 1 To iMaxCol
        Print #1, sTmp
        sTmp = ""
    Next '»For iRow = 1 To iMaxRow

    Print #1, " "
    Print #1, " "
    Print #1, " "
    Print #1, "Words to find:"

    ' *** print the words to look for
    With frmMain.lstInputWord
        lCount = .ListCount - 1
        For lPointer = 0 To lCount
            Print #1, .List(lPointer)
        Next '»For lPointer = 0 To lCount
    End With '»With frmMain.lstInputWord

    Print #1, " "
    Print #1, " "
    Print #1, " "

    ' *** print the key
    For iRow = 1 To iMaxRow
        For iCol = 1 To iMaxCol
            sTmp = sTmp & aryTemp(iRow, iCol) & " "
        Next '»For iCol = 1 To iMaxCol
        Print #1, sTmp
        sTmp = ""
    Next '»For iRow = 1 To iMaxRow

    Print #1, " "
    Print #1, " "
    Print #1, "********** NOTE *****************"
    Print #1, " Import this file to Word and"
    Print #1, " change the font to Courier New."
    Print #1, "*********************************"

    Close #1
    MsgBox "Export done. File name is " & sFilePath, vbInformation + vbOKOnly + vbDefaultButton1, "Export"

End Sub

Function FitWord(sWord As String, iRow As Integer, iCol As Integer, iDirection As Integer) As Boolean
    ' *******************************
    ' Direction       Operation
    ' ---------       ---------
    '     1           Left -> Right
    '     2           Right -> Left
    '     3           Top -> Bottom
    '     4           Bottom -> Top
    '     5           Top Left -> Bottom Right
    '     6           Bottom Right -> Top Left
    '     7           Top Right -> Bottom Left
    '     8           Bottom Right -> Top Left

    Dim iTmp As Integer
    Dim iMaxRow As Integer
    Dim iMaxCol As Integer
    Dim iPointer As Integer
    Dim iCount As Integer
    Dim iWordChrPointer As Integer

    iWordChrPointer = 1
    ' **** there is a main array and a temp array. We will copy the main array into the
    '      temp array. Then we will try to fit the word into the temp array. Also we are
    '      going to check for a cross-over (Two words crossing each other with the same
    '      chars.) If this passes then we will copy the temp array back to the main array.
    ' *************
    '  Copy the main array to the temp
    '
    With frmMain
        iMaxRow = Val(.txtGridRows.Text)
        iMaxCol = Val(.txtGridCol.Text)
        ReDim aryTemp(iMaxRow, iMaxCol) As String
        For iCount = 1 To iMaxRow
            For iPointer = 1 To iMaxCol
                aryTemp(iCount, iPointer) = aryMain(iCount, iPointer)
            Next '»For iPointer = 1 To iMaxCol
        Next '»For iCount = 1 To iMaxRow

        FitWord = False
        Select Case iDirection
            Case 1  ' Left - Right (Cols only)
                ' *** quick check
                If Len(sWord) <= (Val(.txtGridCol.Text) - iCol) Then
                    ' *** lay the word on the grid
                    iCount = (iCol + Len(sWord) - 1)
                    
                    iWordChrPointer = 1
                    FitWord = True
                    For iPointer = iCol To iCount
                        If aryTemp(iRow, iPointer) = "." _
                        Or Asc(aryTemp(iRow, iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                            aryTemp(iRow, iPointer) = Mid(sWord, iWordChrPointer, 1)
                            iWordChrPointer = iWordChrPointer + 1
                        Else
                            FitWord = False
                            Exit Function
                        End If '»If aryTemp(iRow, iPointer) = " " _
                        Or Asc(aryTemp(iRow, iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                    Next '»For iPointer = iCol To iCount

                End If '»If Len(sWord) <= (Val(.txtGridCol.Text) - iCol) Then

            Case 2   ' Right - Left (Cols Only)
                If Len(sWord) <= iCol Then
                    ' *** lay the word on the grid
                    iCount = Len(sWord)
                    iWordChrPointer = 1
                    FitWord = True
                    For iPointer = iCount To 1 Step -1
                        If aryTemp(iRow, iPointer) = "." _
                        Or Asc(aryTemp(iRow, iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                            aryTemp(iRow, iPointer) = Mid(sWord, iWordChrPointer, 1)
                            iWordChrPointer = iWordChrPointer + 1

                        Else
                            FitWord = False
                            Exit Function
                        End If '»If aryTemp(iRow, iPointer) = " " _
                        Or Asc(aryTemp(iRow, iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                    Next '»For iPointer = iCount To 1 Step -1

                    FitWord = True
                End If '»If Len(sWord) <= iCol Then
            Case 3  ' Top to Bottom
                If Len(sWord) <= (Val(.txtGridRows.Text) - iRow) Then
                    FitWord = True

                    ' *** lay the word on the grid
                    iCount = (iRow + Len(sWord) - 1)
                    iWordChrPointer = 1
                    FitWord = True
                    For iPointer = iRow To iCount
                        If aryTemp(iPointer, iCol) = "." _
                        Or Asc(aryTemp(iPointer, iCol)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                            aryTemp(iPointer, iCol) = Mid(sWord, iWordChrPointer, 1)
                            iWordChrPointer = iWordChrPointer + 1
                        Else
                            FitWord = False
                            Exit Function

                        End If '»If aryTemp(iPointer, iCol) = " " _
                        Or Asc(aryTemp(iPointer, iCol)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                    Next '»For iPointer = iRow To iCount


                End If '»If Len(sWord) <= (Val(.txtGridRows.Text) - iRow) Then
            Case 4  ' Bottom to Top
                If Len(sWord) <= iRow Then

                    FitWord = True
                    ' *** lay the word on the grid
                    iCount = Len(sWord)
                    iWordChrPointer = 1
                    FitWord = True
                    For iPointer = iCount To 1 Step -1
                        If aryTemp(iPointer, iCol) = "." _
                        Or Asc(aryTemp(iPointer, iCol)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                            aryTemp(iPointer, iCol) = Mid(sWord, iWordChrPointer, 1)
                            iWordChrPointer = iWordChrPointer + 1
                        Else
                            FitWord = False
                            Exit Function

                        End If '»If aryTemp(iPointer, iCol) = " " _
                        Or Asc(aryTemp(iPointer, iCol)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                    Next '»For iPointer = iCount To 1 Step -1
                End If '»If Len(sWord) <= iRow Then

            Case 5  'Top Left -> Bottom Right
                iTmp = Len(sWord)
                iMaxRow = Val(.txtGridRows.Text)
                iMaxCol = Val(.txtGridCol.Text)
                
                For iPointer = 1 To iTmp
                    If iCol + (iPointer - 1) > iMaxCol _
                      Or iRow + (iPointer - 1) > iMaxRow Then
                        FitWord = False
                        Exit Function
                    Else
                        ' *** lay the word on the array
                        If aryTemp(iRow + (iPointer - 1), iCol + (iPointer - 1)) = "." _
                        Or Asc(aryTemp(iRow + (iPointer - 1), iCol + (iPointer - 1))) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                            aryTemp(iRow + (iPointer - 1), iCol + (iPointer - 1)) = Mid(sWord, iWordChrPointer, 1)
                            iWordChrPointer = iWordChrPointer + 1
                        Else
                           FitWord = False
                           Exit Function
                        End If '»If aryTemp(iRow + iPointer, iCol + iPointer) = " " _
                        Or Asc(aryTemp(iRow + iPointer, iCol + iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then

                    End If '»If iCol + iPointer > iMaxCol _
                      Or iRow + iPointer > iMaxRow Then
                Next '»For iPointer = 0 To iTmp
                FitWord = True

            Case 6  ' Bottom Right -> Top Left
                iTmp = Len(sWord)
                For iPointer = 1 To iTmp
                    If iCol - (iPointer - 1) = 0 _
                      Or iRow - (iPointer - 1) = 0 Then
                        FitWord = False
                        Exit Function
                    Else
                        ' *** lay the word on the array
                        If aryTemp(iRow - (iPointer - 1), iCol - (iPointer - 1)) = "." _
                        Or Asc(aryTemp(iRow - (iPointer - 1), iCol - (iPointer - 1))) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                            aryTemp(iRow - (iPointer - 1), iCol - (iPointer - 1)) = Mid(sWord, iWordChrPointer, 1)
                            iWordChrPointer = iWordChrPointer + 1
                        Else
                           FitWord = False
                           Exit Function
                            
                        End If '»If aryTemp(iRow - (iPointer - 1), iCol - (iPointer - 1)) = " " _
                        Or Asc(aryTemp(iRow - (iPointer - 1), iCol - (iPointer - 1))) = Asc(Mid(sWord, iWordChrPointer, 1)) Then

                    End If '»If iCol - iPointer = 0 _
                      Or iRow - iPointer = 0 Then
                Next '»For iPointer = 0 To iTmp
                FitWord = True

            Case 7
                iTmp = Len(sWord) - 1
                For iPointer = 0 To iTmp
                    If iCol - iPointer = 0 _
                      Or iRow + iPointer > iMaxRow Then
                        FitWord = False
                        Exit Function
                    Else
                        ' *** lay the word on the array
                        If aryTemp(iRow + iPointer, iCol - iPointer) = "." _
                        Or Asc(aryTemp(iRow + iPointer, iCol - iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                            aryTemp(iRow + iPointer, iCol - iPointer) = Mid(sWord, iWordChrPointer, 1)
                            iWordChrPointer = iWordChrPointer + 1
                        Else
                           FitWord = False
                           Exit Function
                            
                        End If '»If aryTemp(iRow + iPointer, iCol - iPointer) = " " _
                        Or Asc(aryTemp(iRow + iPointer, iCol - iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then

                    End If '»If iCol - iPointer = 0 _
                      Or iRow + iPointer > iMaxRow Then
                Next '»For iPointer = 0 To iTmp
                FitWord = True

            Case 8
                iTmp = Len(sWord) - 1
                For iPointer = 0 To iTmp
                    If iCol + iPointer > iMaxCol _
                      Or iRow - iPointer = 0 Then
                        FitWord = False
                        Exit Function
                    Else
                        ' *** lay the word on the array
                        If aryTemp(iRow - iPointer, iCol + iPointer) = "." _
                        Or Asc(aryTemp(iRow - iPointer, iCol + iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then
                            aryTemp(iRow - iPointer, iCol + iPointer) = Mid(sWord, iWordChrPointer, 1)
                            iWordChrPointer = iWordChrPointer + 1
                        Else
                           FitWord = False
                           Exit Function
                            
                        End If '»If aryTemp(iRow - iPointer, iCol + iPointer) = " " _
                        Or Asc(aryTemp(iRow - iPointer, iCol + iPointer)) = Asc(Mid(sWord, iWordChrPointer, 1)) Then

                    End If '»If iCol + iPointer > iMaxCol _
                      Or iRow - iPointer = 0 Then
                Next '»For iPointer = 0 To iTmp
                FitWord = True
        End Select '»Select Case iDirection

        For iCount = 1 To iMaxRow
            For iPointer = 1 To iMaxCol
                aryMain(iCount, iPointer) = aryTemp(iCount, iPointer)
                'If Asc(aryMain(iCount, iPointer)) > 32 Then Stop
            Next '»For iPointer = 1 To iMaxCol
        Next '»For iCount = 1 To iMaxRow

    End With '»With frmMain
End Function

Sub FitWordOnGrid()
    Dim lPointer As Long
    Dim lCount As Long
    Dim lMaxTryCount As Long
    Dim lCurrentTryCount As Long
    Dim lWordCount As Long
    Dim lCouldNotFitWordFlg As Boolean
    Dim iTmpRow As Integer
    Dim iTmpCol As Integer
    Dim iDir As Integer
    '
    ' *** try to fit the word on the grid this many times
    lMaxTryCount = 1000
    '
    '
    lCouldNotFitWordFlg = False
    With frmMain
        ' *** how many words on the list
        lCount = UBound(aryListBox, 1) - 1
        
        ' *** start looping through the list
        iTmpRow = 0
        iTmpCol = 0
        iDir = 0

        ' *** loop through the array
        '     Remember that the words on the list were resorted by their size.
        '     This is very important for it's easer to fit the longer words first.
        '
        For lPointer = 1 To lCount
            ' *** check the max word number
            If lPointer > Val(.txtWordCount.Text) Then Exit Sub
            For lCurrentTryCount = 0 To lMaxTryCount

                ' *** get a randum row, col, and direction
                Call RandPos(iTmpRow, iTmpCol, iDir)

                ' *** check to see if these numbers work in fitting the word
                If FitWord(aryListBox(lPointer), iTmpRow, iTmpCol, iDir) = True Then
                    lCouldNotFitWordFlg = False  ' we can fit the word
                    Exit For
                End If '»If FitWord(aryListBox(lPointer), iTmpRow, iTmpCol, iDir) = True Then
            Next '»For lCurrentTryCount = 0 To lMaxTryCount
        Next '»For lPointer = 1 To lCount
    End With '»With frmMain
End Sub

Sub ListBoxControl(sMode As String)
    ' ***********************************************
    '    MODE    OPERATION
    '    ----    ---------
    '    SAVE    Save the new word
    '    CLEAR   Clear the list box
    '    DELETE  Remove the selected word
    '    SORT    Sort the list from Longest to Shortest word
    ' ************************************************
    Dim lPointer As Long
    Dim lCount As Long
    
    
    With frmMain.lstInputWord
        Select Case UCase(sMode)
            Case "SORT"  ' by the longest word to the shortest
                Dim sLastWord As String
                Dim bFlg As Boolean
                Dim lWordCount As Long
                Dim lUpper As Long
                Dim bSortLoopFlg As Boolean
                
                lCount = .ListCount - 1
                ' *** set the item data to 0
                '     this will be used as a flag later
                
                For lPointer = 0 To lCount
                   .ItemData(lPointer) = 0
                Next
                ReDim aryListBox(0)
                bSortLoopFlg = False
                Do Until bSortLoopFlg = True
                bFlg = True
                    bSortLoopFlg = True
                    For lPointer = 0 To lCount
                        If bFlg = True Then
                            ' *** get the first word that the item data is zero
                            If .ItemData(lPointer) <> -1 Then
                                bSortLoopFlg = False
                                ' *** get the first word
                                sLastWord = .List(lPointer)
                                lWordCount = lPointer
                                bFlg = False
                            End If '»If .ItemData(lPointer) <> -1 Then
                        Else
                            ' *** get the next word that the item data is zero
                            If .ItemData(lPointer) <> -1 Then
                                ' *** we still have words to check
                                bSortLoopFlg = False
                                ' *** find the largest
                                If Len(Trim(sLastWord)) < Len(Trim(.List(lPointer))) Then
                                    sLastWord = .List(lPointer)
                                    lWordCount = lPointer
                                End If '»If Len(sLastWord) < .List(lPointer) Then

                            End If '»If .ItemData(lPointer) <> -1 Then
                        End If '»If bFlg = True Then
                    Next '»For lPointer = 0 To lCount
                    lUpper = UBound(aryListBox, 1) + 1

                    ReDim Preserve aryListBox(lUpper)
                    aryListBox(lUpper) = sLastWord
                    .ItemData(lWordCount) = -1
                Loop '»Do Until bSortLoopFlg = True

            Case "DELETE"
                If .SelCount = 0 Then Exit Sub
                lCount = .ListCount - 1
                For lPointer = 0 To lCount
                    If .Selected(lPointer) = True Then
                        .RemoveItem lPointer
                        Exit Sub
                    End If '»If .Selected(lPointer) = True Then
                Next '»For lPointer = 0 To lCount
                
            Case "SAVE"
                .AddItem UCase(frmMain.txtInputWord.Text)
                .ItemData(.NewIndex) = 0
                frmMain.txtInputWord.Text = ""
                frmMain.txtInputWord.SetFocus
                
            Case "CLEAR"
                .Clear
                frmMain.txtWord.Text = ""
                frmMain.txtInputWord.Text = ""
                frmMain.cmdStart.Enabled = False
                frmMain.cmdAuto.Enabled = False
                frmMain.cmdFinish.Enabled = False
                frmMain.cmdExport.Enabled = False
                
                frmMain.txtInputWord.SetFocus
        End Select '»Select Case UCase(sMode)
    End With '»With frmMain.lstInputWord
End Sub

Sub RandPos(iRow As Integer, iCol As Integer, iDirection As Integer)
    ' *******************************
    ' Direction       Operation
    ' ---------       ---------
    '     1           Left -> Right
    '     2           Right -> Left
    '     3           Top -> Bottom
    '     4           Bottom -> Top
    '     5           Top Left -> Bottom Right
    '     6           Bottom Right -> Top Left
    '     7           Top Right -> Bottom Left
    '     8           Bottom Right -> Top Left

    With frmMain

        Randomize
        ' *** get the row
        iRow = Int((.txtGridRows * Rnd) + 1)
        iCol = Int((.txtGridCol * Rnd) + 1)
        
        ' **** set the easy/hard rnd numbers
        With frmMain.optEasy
            If .Value = False Then
                iDirection = Int((8 * Rnd) + 1)   ' Generate random value between 1 and 8
            Else
                iDirection = Int((4 * Rnd) + 1)   ' Generate random value between 1 and 4
            End If '»If .Value = False Then
        End With '»With frmMain.optEasy

    End With '»With frmMain
End Sub



Sub ShowFinishedWord()
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    Dim lRowPointer As Long
    Dim lColPointer As Long
    Dim sCTmp As String
    Dim sRTmp As String


    With frmMain
    
        .txtWord.Text = ""
        lMaxRow = Val(.txtGridRows.Text)
        lMaxCol = Val(.txtGridCol.Text)

        ' *** copy the main array to the temp array. Durring the print command we will
        '     print the puzzle with the dots (solution view) and the fill in view.
        For lRowPointer = 1 To lMaxRow
            For lColPointer = 1 To lMaxCol
                aryTemp(lRowPointer, lColPointer) = aryMain(lRowPointer, lColPointer)
            Next '»For lColPointer = 1 To lMaxCol
        Next '»For lRowPointer = 1 To lMaxRow


        For lRowPointer = 1 To lMaxRow
            For lColPointer = 1 To lMaxCol
                ' *** fill in the dots with random letters
                If aryMain(lRowPointer, lColPointer) = "." Then
                    aryMain(lRowPointer, lColPointer) = Chr(Int((26 * Rnd) + 65))
                End If '»If aryMain(lRowPointer, lColPointer) = "." Then
                sCTmp = sCTmp & aryMain(lRowPointer, lColPointer) & " "
                ' If Asc(aryMain(lRowPointer, lColPointer)) <> 46 Then Stop
            Next '»For lColPointer = 1 To lMaxCol
            sRTmp = sRTmp & sCTmp & vbCrLf
            sCTmp = ""
        Next '»For lRowPointer = 1 To lMaxRow
        .txtWord.Text = sRTmp
    End With '»With frmMain
End Sub

Sub ShowWordOnWordText()
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    Dim lRowPointer As Long
    Dim lColPointer As Long
    Dim sCTmp As String
    Dim sRTmp As String

    With frmMain
        .txtWord.Text = ""
        lMaxRow = Val(.txtGridRows.Text)
        lMaxCol = Val(.txtGridCol.Text)
        For lRowPointer = 1 To lMaxRow
            For lColPointer = 1 To lMaxCol
               sCTmp = sCTmp & aryMain(lRowPointer, lColPointer) & " "
              ' If Asc(aryMain(lRowPointer, lColPointer)) <> 46 Then Stop
            Next '»For lColPointer = 1 To lMaxCol
            sRTmp = sRTmp & sCTmp & vbCrLf
            sCTmp = ""
        Next '»For lRowPointer = 1 To lMaxRow
        .txtWord.Text = sRTmp
    End With '»With frmMain
End Sub

Private Sub cmdAuto_Click()

    ' *** redim the arrays
    Dim iCol As Integer
    Dim iColCount As Integer
    Dim iRow As Integer
    Dim iRowCount As Integer
    Dim lAutoCount As Long
    Dim lPointer As Long
    Dim sFilePath As String
    Dim lMaxBarCount As Long
    Dim lBarStep As Long
    Dim sTmpCaption As String
    
    bAutoFrameFlg = False
    With frmMain
        ' *** check for a number value
        If IsNumeric(.txtNumOfPuzzles.Text) = False Then Exit Sub
        
        ' *** get the max units between the two lines
        lMaxBarCount = .linRed.X2 - .linGrn.X1

        ' *** get the number of puzzles that will be built
        lAutoCount = Val(.txtNumOfPuzzles.Text)

        ' *** cal the step
        lBarStep = lMaxBarCount / lAutoCount

        ' *** set the bar to zero
        .linGrn.X2 = .linGrn.X1
        .linRed.X1 = .linGrn.X2

        ' *** save the caption on the label
        sTmpCaption = "Welcome to Word Puzzle by John Nyhart. Currently the program is building puzzles and storing them in "

        Do Until .fraOfPuzzlesToCreate.Width > 6850
            .tmrAuto.Enabled = True
            DoEvents
        Loop '»Do Until .fraOfPuzzlesToCreate.Width > 6850
        .tmrAuto.Enabled = False
    End With '»With frmMain
    Screen.MousePointer = vbHourglass
    For lPointer = 1 To lAutoCount

        sFilePath = App.Path & "WordPuz" & lPointer & ".txt"

        With frmMain
            ' *** change the caption on the label in the frame
            .lblNumOfPuzzles.Caption = sTmpCaption & " " & sFilePath
            DoEvents

            ' *** check the list box for words
            If .lstInputWord.ListCount = 0 Then Exit Sub

            ' *** load the default (dots) in the main array
            iRow = Val(.txtGridRows.Text)
            iCol = Val(.txtGridCol.Text)
            ReDim aryTemp(iRow, iCol) As String
            ReDim aryMain(iRow, iCol) As String
            For iRowCount = 0 To iRow
                For iColCount = 0 To iCol
                    aryMain(iRowCount, iColCount) = "."
                Next '»For iColCount = 0 To iCol
            Next '»For iRowCount = 0 To iRow

            .txtWord.Text = ""
            Call ListBoxControl("SORT")
            Call FitWordOnGrid
            Call ShowWordOnWordText
            Call ShowFinishedWord
            Call AutoExportPuzzle(sFilePath)

            ' *** move the pBar
            .linGrn.X2 = .linGrn.X1 + (lPointer * lBarStep)
            .linRed.X1 = .linGrn.X2
            DoEvents
        End With '»With frmMain
    Next '»For lPointer = 1 To lAutoCount
    With frmMain
        ' *** shrink the frame
        bAutoFrameFlg = True
        Do Until .fraOfPuzzlesToCreate.Width < 1830
            .tmrAuto.Enabled = True
            DoEvents
        Loop '»Do Until .fraOfPuzzlesToCreate.Width < 1830
        .tmrAuto.Enabled = False
    End With '»With frmMain
    Screen.MousePointer = vbNormal
    Call ListBoxControl("CLEAR")
End Sub

Private Sub cmdExit_Click()
End

End Sub

Private Sub cmdExport_Click()
title = InputBox("Please enter a title for your puzzle.", "Enter a title.")


Form1.Show 0
End Sub

Private Sub cmdFinish_Click()
    With frmMain
        .cmdFinish.Enabled = False
        Call ShowFinishedWord
        .cmdExport.Enabled = True

    End With '»With frmMain
End Sub

Private Sub cmdNew_Click()
    Call ListBoxControl("CLEAR")
End Sub

Private Sub cmdRemove_Click()
Call ListBoxControl("DELETE")
End Sub

Private Sub cmdStart_Click()
    ' *** redim the arrays
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iRowCount As Integer
    Dim iColCount As Integer
    
    With frmMain
        ' *** check the list box for words
        If .lstInputWord.ListCount = 0 Then Exit Sub
        iRow = Val(.txtGridRows.Text)
        iCol = Val(.txtGridCol.Text)
        ReDim aryTemp(iRow, iCol) As String
        ReDim aryMain(iRow, iCol) As String
        For iRowCount = 0 To iRow
            For iColCount = 0 To iCol
                aryMain(iRowCount, iColCount) = "."
            Next '»For iColCount = 0 To iCol
        Next '»For iRowCount = 0 To iRow

    Call ListBoxControl("SORT")
    Call FitWordOnGrid
    Call ShowWordOnWordText
    .cmdFinish.Enabled = True
    End With '»With frmMain
End Sub

Private Sub Form_Load()

Call CenterForm(frmMain)
bAutoFrameFlg = False

' *** set the word count to zero
glWordCount = 0

End Sub

Private Sub lstInputWord_Click()
    With frmMain.lstInputWord

        If .SelCount > 0 Then
            frmMain.cmdRemove.Enabled = True
        Else
            frmMain.cmdRemove.Enabled = False
        End If '»If .SelCount > 0 Then

    End With '»With frmMain.lstInputWord
End Sub

Private Sub mnuAbout_Click()
   frmAbout.Show vbModal, Me
End Sub

Private Sub tmrAuto_Timer()
    With frmMain.fraOfPuzzlesToCreate
        If bAutoFrameFlg = False Then     'close -> open

            If .Width < 6855 Then
                .Width = .Width + 200
            End If '»If .Width < 6855 Then

        Else                              ' open -> close

            If .Width > 1815 Then
                .Width = .Width - 200
            End If '»If .Width > 1815 Then

        End If '»If bAutoFrameFlg = False Then
    End With '»With frmMain.fraOfPuzzlesToCreate
End Sub

Private Sub txtInputWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        With frmMain
            If Len(Trim(.txtInputWord.Text)) = 0 Then Exit Sub
            
            .lstInputWord.AddItem UCase(Trim(.txtInputWord.Text))
            .lstInputWord.ItemData(.lstInputWord.NewIndex) = glWordCount
            glWordCount = glWordCount + 1
            .txtInputWord.Text = ""
            .cmdStart.Enabled = True
            .cmdAuto.Enabled = True
            .txtInputWord.SetFocus
        End With '»With frmMain

    End If '»If KeyAscii = 13 Then

End Sub

Private Sub txtWordCount_Change()
    Dim sTmp As String

    With frmMain.txtWordCount
        sTmp = Trim(.Text)

    End With '»With frmMain.txtWordCount
End Sub


