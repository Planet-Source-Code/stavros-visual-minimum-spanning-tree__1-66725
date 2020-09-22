VERSION 5.00
Begin VB.Form frmMST 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Minimum Spanning Tree"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmMST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++----------------------------------------------------------------------------
'//   frmMST
'//     author: Stavros Sirigos
'++----------------------------------------------------------------------------
'//   - Simple Implementation of Prim's Algorithm for computing a
'//     Minimum Spanning Tree (MST), (done mainly for demonstration purposes).
'//
'//   - Can be modified to work efficiently with sparse graphs with the addition
'//     of an adjacency list and a binary heap (see my earlier submission on
'//     Dijkstra's Algorithm on how to accomplish that).
'//
'//   - For the euclidean case, running time can be dramatically reduced by
'//     computing the MST by using only the edges defined by a Delaunay Triangulation
'//     of the vertices. (see Wikipedia for more info on all the above).
'++----------------------------------------------------------------------------
'
'Some bits from Wikipedia:...
'
'http://en.wikipedia.org/wiki/Minimum_spanning_tree
'++------------------------------------------------
'Given a connected, undirected graph, a spanning tree of that graph is a subgraph which
'is a tree and connects all the vertices together. A single graph can have many different
'spanning trees. We can also assign a weight to each edge, which is a number representing
'how unfavorable it is, and use this to assign a weight to a spanning tree by computing
'the sum of the weights of the edges in that spanning tree.
'A minimum spanning tree or minimum weight spanning tree is then a spanning tree with
'weight less than or equal to the weight of every other spanning tree.
'
'One example would be a cable TV company laying cable to a new neighborhood.
'If it is constrained to bury the cable only along certain paths,
'then there would be a graph representing which points are connected by those paths.
'Some of those paths might be more expensive, because they are longer, or require the
'cable to be buried deeper; these paths would be represented by edges with larger weights.
'A spanning tree for that graph would be a subset of those paths that has no cycles
'but still connects to every house. There might be several spanning trees possible.
'A minimum spanning tree would be one with the lowest total cost.
'
'The first algorithm for finding a minimum spanning tree was developed by Czech scientist
'Otakar Boruvka in 1926. Its purpose was an efficient electrical coverage of Bohemia.
'There are now two algorithms commonly used, Prim 's algorithm and Kruskal's algorithm.

'http://en.wikipedia.org/wiki/Prim%27s_algorithm
'++---------------------------------------------
'Prim 's algorithm works as follows:
'
'    * create a tree containing a single vertex, chosen arbitrarily from the graph
'    * create a set containing all the edges in the graph
'    * loop until every edge in the set connects two vertices in the tree
'          o remove from the set an edge with minimum weight that connects a vertex
'            in the tree with a vertex not in the tree
'          o add that edge to the tree

Option Explicit

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private lMvNode As Long     'Current node used in mouse events
Private Pts()   As POINTAPI 'Our collection of 2D vertices. MST must connect those.

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SP Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'++----------------------------------------------------------------------------
'//   PRIM'S Algorithm for the MST. Simple n^2 implementation for Demo purposes.
'++----------------------------------------------------------------------------
Private Function SolveMST(Optional ByVal lDemo As Long) As Single
    
  Dim i         As Long
  Dim j         As Long
  Dim lCur      As Long     'Current examined node in the queue.
  Dim lPos      As Long     'Position in the queue of the next vertex to be connect to the tree.
  Dim lInQ      As Long     'Number of vertices not optimally connected to the tree.
  Dim lNodes    As Long     'Total number of vertices.
  Dim lBestOutV As Long     'ID of the next vertex to be connect to the tree.
  Dim lResult() As Long     'Final edges used in the MST
  Dim lEdges()  As Long     'Final and pending edges used in the MST.
                            'Index means source vertex and value means destination vertex.
                            'e.g. lEdges(4)=2 means that we use the edge (line) from 2 to 4.
  Dim lQueue()  As Long     'List of pending vertices to be connected to the tree.
  Dim sBestCost As Single   'Used for determining closest unconnected vertex to the tree.
  Dim sDist()   As Single   'Distances of vertices to the tree (final or pending).
  Dim sX()      As Single   'Local array of X coordinates
  Dim sY()      As Single   'Local array of X coordinates
  Dim sCurDist  As Single   'Temp variable
  Const sINF    As Single = 1E+38 ' "Infinity"
    
    If IsDimed(Pts) Then
        lNodes = UBound(Pts)
    Else
        MsgBox "Add some nodes first!", vbInformation, "Visual MST"
        Exit Function
    End If

    ReDim lResult(1 To lNodes)
    ReDim lQueue(1 To lNodes)
    ReDim lEdges(1 To lNodes)
    ReDim sDist(1 To lNodes)
    ReDim sX(1 To lNodes)
    ReDim sY(1 To lNodes)
    
    'Initialization - Insert vertex 1 to the tree & compute initial distances
    For i = 1 To lNodes
        lQueue(i) = i
        sX(i) = CDbl(Pts(i).X)
        sY(i) = CDbl(Pts(i).Y)
        sDist(i) = ((sX(1) - sX(i)) * (sX(1) - sX(i)) + _
                    (sY(1) - sY(i)) * (sY(1) - sY(i)))
        lEdges(i) = 1
    Next i
    
    lInQ = lNodes
    lPos = 1
    
    'Main Loop
    For i = 1 To lNodes
    
        lBestOutV = 0
        sBestCost = sINF
        
        'Find the unconnected vertex closest to the tree
        For j = 1 To lInQ
            If sDist(lQueue(j)) < sBestCost Then
                sBestCost = sDist(lQueue(j))
                lBestOutV = lQueue(j)
                lPos = j
            End If
        Next j
        
        If lBestOutV Then
            'Connect the closest vertex to the tree & remove it from pending list.
            lResult(lBestOutV) = lEdges(lBestOutV)
            lQueue(lPos) = lQueue(lInQ)
            lInQ = lInQ - 1
            
            'See if the remaining unconnected vertices are closer to the newly added
            'vertex than their previous closest vertex in the tree.
            For j = 1 To lInQ
                lCur = lQueue(j)
                sCurDist = (sX(lCur) - sX(lBestOutV)) * (sX(lCur) - sX(lBestOutV)) + _
                           (sY(lCur) - sY(lBestOutV)) * (sY(lCur) - sY(lBestOutV))
                
                If sCurDist < sDist(lCur) Then ' It is closer than previous
                    sDist(lCur) = sCurDist          ' Update distance from tree
                    lEdges(lCur) = lBestOutV        ' Update possible connection edge
                End If
            Next j
        End If
        
        SolveMST = SolveMST + Sqr(sDist(lBestOutV)) ' Total length of MST
        
        'Just for demonstrating the progress of the algorithm
        '(triggered by pressing "space" in the GUI)
        If lDemo Then
            If lDemo > 0 Then
                DrawMST lEdges, SolveMST
            Else
                DrawMST lResult, SolveMST
            End If
            Me.Refresh
            Sleep Abs(lDemo)
        End If
    Next i
        
    DrawMST lResult, SolveMST
        
End Function

'++----------------------------------------------------------------------------
'//   Helper function to see if array is dimensioned.
'++----------------------------------------------------------------------------
Private Function IsDimed(Arr() As POINTAPI) As Boolean
    On Error GoTo errHandle
    IsDimed = UBound(Arr) = UBound(Arr)
errHandle:
End Function

'++----------------------------------------------------------------------------
'//   Hard-Coded icon for Nodes.
'//   I wanted to have only 1 source file, with no .frx or .bmp etc., so one
'//   may open it directly from WinZip with no problems.
'++----------------------------------------------------------------------------
Private Sub DrawNodeIcon()
  Dim i As Long
    i = Me.hdc
    SP i, 0, 0, 2533330: SP i, 1, 0, 4305625:  SP i, 2, 0, 3518166:  SP i, 3, 0, 3124181:  SP i, 4, 0, 4763346
    SP i, 0, 1, 5158616: SP i, 1, 1, 14347247: SP i, 2, 1, 16777215: SP i, 3, 1, 13954805: SP i, 4, 1, 1478846
    SP i, 0, 2, 2469079: SP i, 1, 2, 7657198:  SP i, 2, 2, 8447487:  SP i, 3, 2, 7198193:  SP i, 4, 2, 1083319
    SP i, 0, 3, 1486297: SP i, 1, 3, 2933227:  SP i, 2, 3, 2800094:  SP i, 3, 3, 2736362:  SP i, 4, 3, 425903
    SP i, 0, 4, 161705:  SP i, 1, 4, 161705:   SP i, 2, 4, 226469:   SP i, 3, 4, 226984:   SP i, 4, 4, 159906
End Sub

'++----------------------------------------------------------------------------
'//   Simple Info about the controls and the MST results.
'++----------------------------------------------------------------------------
Private Sub TxtOut(ByVal sRet As Single)
  Dim lNodes As Long
    Me.CurrentX = 8
    Me.CurrentY = 0
    If IsDimed(Pts) Then lNodes = UBound(Pts)
    Print "Left MB: Add Node  |  Right MB: Move Node  |  Middle MB: Clear  |  Space: Demo  |  Esc: Exit  |  Current MST length = " & FormatNumber(sRet, 2) & "  |  Nodes = " & lNodes
End Sub

'++----------------------------------------------------------------------------
'//   Drawing of MST etc. on screen
'++----------------------------------------------------------------------------
Private Sub DrawMST(ByRef lResult() As Long, ByVal sRet As Single)

  Dim pt1       As POINTAPI
  Dim i         As Long
  Dim lhDC      As Long
  Dim lPoints   As Long
  Dim lPen      As Long
    
    If IsDimed(Pts) Then
        Me.Cls
        TxtOut sRet
        DrawNodeIcon
        lPoints = UBound(Pts)

        lhDC = Me.hdc
        lPen = CreatePen(vbDash, 1, RGB(160, 160, 160))
        DeleteObject SelectObject(lhDC, lPen)
    
        For i = 1 To lPoints
            If lResult(i) Then ' We found an edge - draw it
                MoveToEx lhDC, Pts(i).X, Pts(i).Y, pt1
                LineTo lhDC, Pts(lResult(i)).X, Pts(lResult(i)).Y
            End If
        Next i
        For i = 1 To lPoints
            BitBlt lhDC, Pts(i).X - 2, Pts(i).Y - 2, 5, 5, hdc, 0, 0, vbSrcCopy
        Next i
        BitBlt lhDC, 0, 0, 5, 5, hdc, 0, 0, vbMergePaint ' Delete our "source" icon
        DeleteObject lPen
    End If
        
End Sub

'++----------------------------------------------------------------------------
'//   Simple method for finding the closest node to a point (used in mouse events).
'++----------------------------------------------------------------------------
Private Function FindClosestNode(ByVal X As Single, ByVal Y As Single) As Long

  Dim i             As Long
  Dim lBest         As Long
  Dim sBest         As Single

    If IsDimed(Pts) Then
        sBest = 1E+38
        For i = 1 To UBound(Pts)
            With Pts(i)
                If (.X - X) * (.X - X) + (.Y - Y) * (.Y - Y) < sBest Then
                    sBest = (.X - X) * (.X - X) + (.Y - Y) * (.Y - Y)
                    lBest = i
                End If
            End With
        Next i
        FindClosestNode = lBest
    End If
            
End Function

'++----------------------------------------------------------------------------
'//   Form Events
'++----------------------------------------------------------------------------
Private Sub Form_Load()
    ShowAbout
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    TxtOut 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
       
    lMvNode = 0
    If Button = vbLeftButton Then
        If Y > 2 * Me.Font.Size Then
            ' Add new node
            If IsDimed(Pts) Then
                ReDim Preserve Pts(1 To UBound(Pts) + 1)
            Else
                ReDim Pts(1 To 1)
            End If
            lMvNode = UBound(Pts)
            Pts(lMvNode).X = CLng(X)
            Pts(lMvNode).Y = CLng(Y)
            
            SolveMST
        End If
        
    ElseIf Button = vbRightButton Then
        'Start moving the closest node
        lMvNode = FindClosestNode(X, Y)
        
    ElseIf Button = vbMiddleButton Then
        'Clear all
        Erase Pts
        Me.Cls
        TxtOut 0
        
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If lMvNode Then
        If Button > 0 Then
            If IsDimed(Pts) Then
                If Y > 2 * Me.Font.Size Then
                    'Move the node and re-solve the MST
                    Pts(lMvNode).X = CLng(X)
                    Pts(lMvNode).Y = CLng(Y)
                    SolveMST
                    Me.Refresh
                End If
            End If
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

  Static lMode As Long
  Static lMsDemo As Long ' Sleep milliseconds (could put buttons for increase/decrease this)
    
    If KeyAscii = vbKeySpace Then
        'Run the algorithm in Slow motion, alternating between two modes
        If lMode = 0 Then lMode = 1
        If lMsDemo = 0 Then lMsDemo = 50
        lMode = lMode * -1
        SolveMST lMode * lMsDemo
                
    ElseIf KeyAscii = vbKeyEscape Then 'Bye!
        Unload Me
        
    End If
    
End Sub

'++----------------------------------------------------------------------------
'//   SHOWS ABOUT MESSAGEBOX
'++----------------------------------------------------------------------------
Private Sub ShowAbout()
    MsgBox "Simple implementation of Prim's algorithm for computing a Minimum Spanning Tree." + vbLf + vbLf _
           + "by Stavros Sirigos. <ssirig@uth.gr>" + vbLf + vbLf _
           + "Quick instructions:" + vbLf + vbLf _
           + "- Left Mouse Button (and Mouse Move): add and optionally move a new node" + vbLf + vbLf _
           + "- Right Mouse Button and Mouse Move: grab and move a node to a new position" + vbLf + vbLf _
           + "- Middle Mouse Button: Clear all" + vbLf + vbLf _
           + "- [Space]: demo - slow motion (2 alternating modes)" + vbLf + vbLf _
           + "- [Escape]: exit" + vbLf + vbLf _
           + "Have fun!", vbInformation, "Visual MST"
End Sub
