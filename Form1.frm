VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text Sort"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "QuickSort"
      Height          =   480
      Left            =   3090
      TabIndex        =   3
      Top             =   195
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sort"
      Height          =   480
      Left            =   1695
      TabIndex        =   2
      Top             =   195
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generate"
      Height          =   480
      Left            =   300
      TabIndex        =   1
      Top             =   195
      Width           =   1185
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   780
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Coded by : Safo
' Email : safo@zoznam.sk
' Program : Text Sort
' Program run faster when it is compile.

Dim Q(1 To 100) As String ' text array

Private Sub Command1_Click()
    Dim time As New clsTimer

    time.StartTimer  ' start timer
    Text_QuickSort Q, 1, 100  ' Sorting
    time.StopTimer   ' stop timer
    
    Me.Caption = "Time of sorting : " & time.Elasped & " ms"
    
    List1.Clear
    
    For i = 1 To 100
        List1.AddItem Q(i)
    Next i
End Sub

Private Sub Command2_Click()
    Dim number As Integer
    
    List1.Clear
    Randomize

    For i = 1 To 100
        number = Rnd() * 7 + 1 ' Random length of text
        Q(i) = ""
        
        For j = 1 To number
            Q(i) = Q(i) & Chr(Rnd() * 25 + 97) ' Random text (a-z)
        Next j
        
        List1.AddItem Q(i)
    Next i
    
End Sub

Private Sub Command3_Click()
    Dim time As New clsTimer

    time.StartTimer  ' start timer
    Text_Sort Q, 1, 100  ' Sorting
    time.StopTimer   ' stop timer
    
    Me.Caption = "Time of sorting : " & time.Elasped & " ms"
    
    List1.Clear
    
    For i = 1 To 100
        List1.AddItem Q(i)
    Next i

End Sub

Private Function Text_Sort(vArray() As String, inLow As Integer, inHi As Integer)
    Dim buf As String
    
    For j = inLow To inHi - 1
    For i = inLow To inHi - 1
    
        If StrComp(vArray(i), vArray(i + 1), vbTextCompare) = 1 Then
            buf = vArray(i)
            vArray(i) = vArray(i + 1)
            vArray(i + 1) = buf
        End If
    
    Next i
    Next j

End Function

Private Function Text_QuickSort(vArray() As String, inLow As Long, inHi As Long)

   Dim pivot   As String
   Dim tmpSwap As String
   Dim tmpLow  As Long
   Dim tmpHi   As Long
    
   tmpLow = inLow
   tmpHi = inHi
   
   pivot = vArray((inLow + inHi) \ 2)
  
   While (tmpLow <= tmpHi)
  
      While (StrComp(vArray(tmpLow), pivot, vbTextCompare) = -1 And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (StrComp(pivot, vArray(tmpHi), vbTextCompare) = -1 And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = vArray(tmpLow)
         vArray(tmpLow) = vArray(tmpHi)
         vArray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then Text_QuickSort vArray, inLow, tmpHi
   If (tmpLow < inHi) Then Text_QuickSort vArray, tmpLow, inHi


End Function
