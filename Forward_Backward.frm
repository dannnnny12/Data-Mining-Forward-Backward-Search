VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   16575
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox Bwd 
      Height          =   4920
      Left            =   11520
      TabIndex        =   15
      Top             =   2880
      Width           =   2775
   End
   Begin VB.ListBox Fwd 
      Height          =   4920
      Left            =   8640
      TabIndex        =   13
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ListBox listUXY 
      Height          =   4920
      Left            =   6480
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ListBox list_content 
      Height          =   5100
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ListBox listHXY 
      Height          =   4920
      Left            =   4440
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ListBox listHX 
      Height          =   4920
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "R76104013_hw1"
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label9 
      Caption         =   "Backward"
      Height          =   255
      Left            =   12480
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Forward"
      Height          =   255
      Left            =   9120
      TabIndex        =   14
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "U(X,Y)"
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "File Content"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "H(X, Y)"
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "H(X)"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Strategy:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "File Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'define attributes
Dim dataset(46, 35), setcount, layer, numerator, denominator As Double
Dim index, attributeValue As Integer
Dim num(46, 34) As Double
Dim raw As Variant, x, couta As Integer
Dim y As Integer
Dim sum As Double, i As Double
Dim HxyArray(36, 36) As Double, data(47, 36), tempU As Double
Dim tempCount(36) As Double
Dim HArray(36) As Double
Dim attributeInd(35) As Integer
Dim col(36) As Double
Dim fullData(47, 36) As Integer
Dim UArray(36, 36) As Double
Dim chosen(35) As Integer


'H(x)
Private Function H(col) As Double
    counta = 0
    '算categorical各值的數量
    For attributeValue = 0 To 35
        For index = 0 To 46
            If fullData(index, col) = attributeValue Then
                counta = counta + 1
            End If
        Next index
        tempCount(attributeValue) = counta
        counta = 0
    Next attributeValue
    Hx = 0
    For j = 0 To 35
        p = tempCount(j) / 47
        Hx = Hx + -p * Log2(p)
    Next j
    HArray(col) = Hx
    H = Hx
End Function
'H(X,Y)
Public Function Hxy(col1, col2)
   counta = 0
    For x = 0 To 35
        For y = 0 To 35
            For index = 0 To 46
                If fullData(index, col1) = x And fullData(index, col2) = y Then
                    counta = counta + 1
                End If
            Next index
            Pxy = counta / 47
            tempHxy = tempHxy + -Pxy * Log2(Pxy)
            counta = 0
        Next y
    Next x
    HxyArray(col1, col2) = tempHxy
    Hxy = HxyArray(col1, col2)
End Function
'U
Private Function U(col1, col2)
    If CLng(H(col1) + H(col2)) = 0 Then
    U = 1
    Else
        U = 2 * ((HArray(col1) + HArray(col2) - HxyArray(col1, col2)) / (HArray(col1) + HArray(col2)))
    End If
    UArray(col1, col2) = U
End Function

Function Goodness()
    numerator = 0
    denominator = 0
    '分子
    For i = 0 To 34
        If chosen(i) = 1 Then
       numerator = numerator + UArray(i, 35)

        End If
    Next i
    '分母
    For i = 0 To 34
        For j = 0 To 34
            If chosen(i) = 1 And chosen(j) = 1 Then
        
            denominator = denominator + UArray(j, i)
            End If
        Next j
    Next i
    
    '計算goodness
    denominator = Sqr(denominator)
    If denominator <> 0 Then
        Goodness = numerator / (denominator)
        
    Else
    Goodness = 0
    End If
End Function


Private Sub Combo1_Click()
    If Combo1.Text = "Search Forward" Then '----------
        Dim i As Integer
        Fwd.Clear
       
        maxGoodness = 0
        
        For i = 0 To 34
            chosen(i) = 0 '初始化，全部設為0
        Next i
    'forward
        For j = 0 To 34
                index = -1 '當index=-1,停止
            For i = 0 To 34
                If chosen(i) <> 1 Then
                    chosen(i) = 1
                    If Goodness > maxGoodness Then
                    maxGoodness = Goodness
                    index = i
                    
                    Fwd.AddItem ("Attribute chosen：A" & i + 1)
                    
                    For b = 0 To 35
                    If chosen(b) = 1 Then
                    Fwd.AddItem ("A" & b + 1)
                    End If
                    Next b
                    
                  Fwd.AddItem ("Goodness：" & maxGoodness)
                  
                    End If
                chosen(i) = 0
                End If
            Next i
            If index = -1 Then Exit For
                chosen(index) = 1
        Next j
    
    Fwd.AddItem ("-------------------")
    
    For i = 0 To 34
        If chosen(i) = 1 Then
            Fwd.AddItem ("Attribute subset A" & i + 1)
        End If
    Next i
    End If
    
   If Combo1.Text = "Search Backward" Then '------------
        maxGoodness = 0
      Bwd.Clear
      
        For i = 0 To 34
            chosen(i) = 1
        Next i
    For n = 0 To 34
        index = -1
        For i = 0 To 34
            If chosen(i) = 1 Then
                chosen(i) = 0
                If Goodness > maxGoodness Then
                    maxGoodness = Goodness
                    index = i
                    
                    
        For w = 0 To 34
            If chosen(w) = 0 Then
                Bwd.AddItem ("A" & w + 1 & "removed")
              
                End If
        Next w
        Bwd.AddItem ("Goodness：" & maxGoodness)
                            
                        End If
                        chosen(i) = 1
                    End If
                Next i
                
                If index = -1 Then Exit For
                    chosen(index) = 0
              Next n
    
    Bwd.AddItem ("-------------------")
    
    For i = 0 To 34
        If chosen(i) = 1 Then
            Bwd.AddItem ("Attribute subset A" & i + 1)
        End If
    Next i
    End If
End Sub

Private Sub Form_Load()
    Combo1.AddItem "Search Forward" 'ListIndex = 0
    Combo1.AddItem "Search Backward" 'ListIndex = 1
     
    Dim i As Integer
    Dim path As String
    i = 0
    Open App.path & "\soybean-small.txt" For Input As #1
    
    
    Do While Not EOF(1)
        Line Input #1, rawTmp
        For j = 0 To 35
        Dim s As Variant
        s = Split(rawTmp, ",")
            If j = 35 Then
                fullData(i, j) = CInt(Split(s(j), "D")(1)) '將最後一欄之D去掉
            Else
                fullData(i, j) = CInt(s(j))
            End If
        Next
        i = i + 1
    Loop
    Close #1
   
    For i = 0 To 46
        For j = 0 To 35
        list_content.AddItem (fullData(i, j))
        Next j
    Next i
    For i = 0 To 35
        listHX.AddItem (H(i))
         For n = 0 To 35
            listHXY.AddItem ("H" & i + 1 & "," & n + 1 & ":" & (Hxy(n, i)))
        Next n
    Next i
    
    For i = 0 To 35
         For n = 0 To 35
            listUXY.AddItem ("U" & i + 1 & "," & n + 1 & ":" & (U(n, i)))
        Next n
    Next i
            
 End Sub
' log function
Static Function Log2(x) As Double
    If x <> 0 Then
        Log2 = Log(x) / Log(2#) '轉為以2為底
    End If
End Function
     

