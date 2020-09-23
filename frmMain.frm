VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convert Number to Words"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumberAsWord 
      Height          =   7455
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   11055
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert To Words"
      Height          =   435
      Left            =   4800
      TabIndex        =   1
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox txtNumber 
      Height          =   300
      Left            =   60
      MaxLength       =   306
      TabIndex        =   0
      Top             =   60
      Width           =   10995
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ConvertNumberToText(vstrNumber As String) As String
   Dim strNumber           As String
   Dim strCluster          As String
   Dim strNumberAsText     As String
   Dim strTempNumber       As String
   Dim strTemp             As String
   Dim strPosition         As String
   Dim blnDone             As Boolean
   Dim intPosition         As Integer
   Dim aryPosition(101)     As String
   Dim aryPrefix(9)        As String
   Dim aryBody(10)        As String
   
   strNumber = vstrNumber
   
      '---Replace any commas with an empty string.
   strNumber = Replace(strNumber, ",", "")
   
      '---Determine if there are any characters other than numbers.  If there are, then
      '---exit the function.
   For intPosition = 1 To Len(strNumber)
      If (Asc(Mid(strNumber, intPosition, 1)) < 48) Or (Asc(Mid(strNumber, intPosition, 1)) > 57) Then
         MsgBox "Please enter a whole number"
         Exit Function
      End If
   Next intPosition
   
      '---Prepare the prefix array.  For every set of ten groups of three, this array is
      '---called once.  For example, from 10^36 to 10^60 (stepping 3 positions each time),
      '---the name would be decillion: UNdecillion (10^36), DUOdecillion (10^39), etc.
   aryPrefix(1) = " "
   aryPrefix(2) = " Un"
   aryPrefix(3) = " Duo"
   aryPrefix(4) = " Tre"
   aryPrefix(5) = " Quattuor"
   aryPrefix(6) = " Quin"
   aryPrefix(7) = " Sex"
   aryPrefix(8) = " Septen"
   aryPrefix(9) = " Octo"
   aryPrefix(0) = " Novem"
   
      '---Prepare the body array.
   aryBody(1) = "Decillion "           '---From 10^33 to 10^60
   aryBody(2) = "Vigintillion "        '---From 10^63 to 10^90
   aryBody(3) = "Trigintillion "       '---From 10^93 to 10^120
   aryBody(4) = "Quadragintillion "    '---From 10^123 to 10^150
   aryBody(5) = "Quinquagintillion "   '---From 10^153 to 10^180
   aryBody(6) = "Sexuagintillion "     '---From 10^183 to 10^210
   aryBody(7) = "Septuagintillion "    '---From 10^213 to 10^240
   aryBody(8) = "Octogintillion "      '---From 10^243 to 10^270
   aryBody(9) = "Nonagintillion "      '---From 10^273 to 10^300
   aryBody(10) = "Centillion "         '---From 10^303 to 10^333
   
      '---First set the basic non-pattern names.
   aryPosition(0) = ""
   aryPosition(1) = " Thousand "
   aryPosition(2) = " Million "
   aryPosition(3) = " Billion "
   aryPosition(4) = " Trillion "
   aryPosition(5) = " Quadrillion "
   aryPosition(6) = " Quintillion "
   aryPosition(7) = " Sextillion "
   aryPosition(8) = " Septillion "
   aryPosition(9) = " Octillion "
   aryPosition(10) = " Nonillion "
   
      '---The rest of the names follow a set pattern up to centillion (10^303).
   For intPosition = 11 To 101

      If (intPosition - 1) = (intPosition \ 10) * 10 Then
         aryPosition(intPosition) = aryPrefix(intPosition Mod 10) & aryBody(Int(intPosition / 10))
      Else
         aryPosition(intPosition) = aryPrefix(intPosition Mod 10) & LCase(aryBody(Int(intPosition / 11)))
      End If
     
   Next intPosition
   
   strPosition = ""
   strNumberAsText = ""
   strNumber = StrReverse(strNumber)
   blnDone = False
   intPosition = 0
      
      '---Starting with the right-most group of three, build the word-number.
   Do While Not blnDone
   
         '---The string was reversed to go from right to left.  Now, get the
         '---first three characters, then reverse the result, which will set
         '---the group of three in the original order.
      strCluster = StrReverse(Mid(strNumber, 1, 3))
      strTempNumber = ""
      
         '---Each iteration, strNumber is trimmed by three until there is nothing left.
      If (Len(strNumber) - 3) < 1 Then
         blnDone = True
      Else
         strNumber = Mid(strNumber, 4)
      End If
      
         '---Pad anything short of three chatacters with 0 (so 53 would be 053).
      Do While Len(strCluster) < 3
         strCluster = "0" & strCluster
      Loop
   
         '---Hundreds
      strTemp = GetTextNumber(Val(Left(strCluster, 1)))
      If Len(strTemp) > 0 Then strTempNumber = strTemp & " Hundred "
      
         '---Tens
      strTemp = GetTextNumber(Val(Mid(strCluster, 2, 2)))
      If Len(strTemp) = 0 Then
         strTemp = GetTextNumber(Val(Mid(strCluster, 2, 1)) * 10)
         strTempNumber = strTempNumber & strTemp
         strTemp = GetTextNumber(Val(Mid(strCluster, 3, 1)))
         If Len(strTemp) > 0 Then
            strTempNumber = strTempNumber & " " & strTemp
         End If
      Else
         strTempNumber = strTempNumber & strTemp
      End If

      If Len(strTempNumber) + 0 Then strNumberAsText = strTempNumber & aryPosition(intPosition) & strNumberAsText

      intPosition = intPosition + 1

   Loop
   
   ConvertNumberToText = strNumberAsText
   
End Function

Private Function GetTextNumber(intNumber As Integer) As String

      '---This function translates the common numbers into words.
   Select Case intNumber
      Case 0
         GetTextNumber = ""
      Case 1
         GetTextNumber = "One"
      Case 2
         GetTextNumber = "Two"
      Case 3
         GetTextNumber = "Three"
      Case 4
         GetTextNumber = "Four"
      Case 5
         GetTextNumber = "Five"
      Case 6
         GetTextNumber = "Six"
      Case 7
         GetTextNumber = "Seven"
      Case 8
         GetTextNumber = "Eight"
      Case 9
         GetTextNumber = "Nine"
      Case 10
         GetTextNumber = "Ten"
      Case 11
         GetTextNumber = "Eleven"
      Case 12
         GetTextNumber = "Twelve"
      Case 13
         GetTextNumber = "Thirteen"
      Case 14
         GetTextNumber = "Fourteen"
      Case 15
         GetTextNumber = "Fifteen"
      Case 16
         GetTextNumber = "Sixteen"
      Case 17
         GetTextNumber = "Seventeen"
      Case 18
         GetTextNumber = "Eighteen"
      Case 19
         GetTextNumber = "Nineteen"
      Case 10
         GetTextNumber = "Ten"
      Case 20
         GetTextNumber = "Twenty"
      Case 30
         GetTextNumber = "Thirty"
      Case 40
         GetTextNumber = "Fourty"
      Case 50
         GetTextNumber = "Fifty"
      Case 60
         GetTextNumber = "Sixty"
      Case 70
         GetTextNumber = "Seventy"
      Case 80
         GetTextNumber = "Eighty"
      Case 90
         GetTextNumber = "Ninety"
   End Select
   
End Function

Private Sub cmdConvert_Click()
   txtNumberAsWord.Text = ConvertNumberToText(txtNumber.Text)
End Sub
