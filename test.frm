VERSION 5.00
Object = "{766D1D78-6698-11D3-8D23-02608C44B837}#1.0#0"; "SUPFILL.OCX"
Begin VB.Form test 
   Caption         =   "SuperFill OCX - Test"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   StartUpPosition =   3  'Windows Default
   Begin AxSuperFill.SuperFill SuperFill1 
      Left            =   660
      Top             =   3960
      _ExtentX        =   1111
      _ExtentY        =   1164
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fill Array and draw"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   4980
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pattern Fill"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   4980
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solid Fill"
      Height          =   495
      Left            =   1380
      TabIndex        =   1
      Top             =   4980
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   300
      TabIndex        =   0
      Top             =   4980
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the form for draw a polygon and try the buttons"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6675
   End
End
Attribute VB_Name = "test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MouseOn As Boolean
Dim Sx As Single, Sy As Single, Bx As Single, By As Single

Dim Px() As Double
Dim Py() As Double

Dim nVert As Integer


Sub AddPoly(x As Double, y As Double)
    
    nVert = nVert + 1
    
    ReDim Preserve Px(nVert)
    ReDim Preserve Py(nVert)
    
    Px(nVert) = x
    Py(nVert) = y
    
End Sub

Private Sub Command1_Click()
  Cls
  nVert = -1
  
End Sub

Private Sub Command2_Click()
  
  Dim RetArray() As Double, NL As Long
  
  SuperFill1.FillMode = 0 ' 0- Output on Object, 1- Get Array of lines
  SuperFill1.FillColor = QBColor(14)
  SuperFill1.FillPattern = ""
  SuperFill1.FillScale = 10
  
  
' Note:.  Px() e Py() are vertexes of polygon
'         with base = 0  (first vertex on element 0)

  
  SuperFill1.SuperFill Me, Px, Py, nVert, RetArray, NL
  
End Sub

Private Sub Command3_Click()
  Dim RetArray() As Double, NL As Long
  
  SuperFill1.FillMode = 0
  SuperFill1.FillColor = QBColor(14)
  SuperFill1.FillPattern = "onde.mtp"
  SuperFill1.FillScale = 60
  
 ' AutoRedraw = True
  
  SuperFill1.SuperFill Me, Px, Py, nVert, RetArray, NL


End Sub

Private Sub Command4_Click()
  Dim RetArray() As Double, NL As Long
  Dim i As Integer
  Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
  
  ' FillMode 1 can used with solid
  ' RetArray will have n horizontal lines for each y of polygon
  
  SuperFill1.FillMode = 1
  SuperFill1.FillPattern = "ghiaia.mtp"
  SuperFill1.FillScale = 50
  
  SuperFill1.SuperFill Me, Px, Py, nVert, RetArray, NL

  For i = 1 To NL Step 4
      x1 = RetArray(i)
      y1 = RetArray(i + 1)
      x2 = RetArray(i + 2)
      y2 = RetArray(i + 3)
      Line (x1, y1)-(x2, y2), QBColor(0)
  Next
  
End Sub

Private Sub Form_Load()
  
  
  ChDir App.Path
  nVert = -1
 
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Not MouseOn Then
     MouseOn = True
  End If
  
     If Button = 1 Then
       
       ' Memorizazione Verici
       
         AddPoly CDbl(x), CDbl(y)
         
         If nVert > 0 Then
            Line (Px(nVert - 1), Py(nVert - 1))-(Px(nVert), Py(nVert))
         End If
     
     Else
        
       ' Memorizazione Ultimo Vertice del poligono
        
        AddPoly CDbl(x), CDbl(y)
        If nVert > 1 Then
            Line (Px(0), Py(0))-(Px(nVert), Py(nVert))
        End If
        
        MouseOn = False
     
     End If
  
 
  Sx = x
  Sy = y
  Bx = x
  By = y



End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 
 ' Animazione Linee
 
 Dim BC As Long
 If MouseOn Then
    BC = ForeColor
    ForeColor = BackColor
    DrawMode = 6
    
    Line (Sx, Sy)-(Bx, By)
    Line (Sx, Sy)-(x, y)
    

    ForeColor = BC
    DrawMode = 13
    
    Bx = x
    By = y
 End If


End Sub




