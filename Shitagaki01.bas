Attribute VB_Name = "Shitagaki01"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 分解された引出線の疑似的復元プログラム
'
' 新たに引出線を作成し重なる線オブジェクトおよび矢印ブロックを削除する
'------------------------------------------------------------------------------
Public Sub RestoreLeader()
    
    On Error GoTo Error_Handler
    
    ' 引出線作図用の2点を指定
    Dim firstPoint As Variant
    Dim secondPoint As Variant
    firstPoint = ThisDrawing.Utility.GetPoint(, "引出線の1点目を指定 [Cancel(ESC)]")
    secondPoint = ThisDrawing.Utility.GetPoint(firstPoint, "引出線の2点目を指定 [Cancel(ESC)]")
    
    Dim leaderPoint(5) As Double
    Dim i As Long
    For i = 0 To 2
        leaderPoint(i) = firstPoint(i)
        leaderPoint(i + 3) = secondPoint(i)
    Next i
    
    ' 新しい引出線を作図
    Dim newLeader As AcadLeader
    Dim leaderType As Integer
    Dim annotationObject As AcadObject
    leaderType = acLineWithArrow
    Set annotationObject = Nothing
    
    Set newLeader = ThisDrawing.ModelSpace.AddLeader(leaderPoint, annotationObject, leaderType)
    
    ' 引出線作図の2点を利用して選択セットを作成
    Dim newSelectionSet As AcadSelectionSet
    Set newSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    
    Dim selectMode As Integer
    selectMode = acSelectionSetWindow
    newSelectionSet.Select selectMode, firstPoint, secondPoint
    
    ' 新しい引出線の単位ベクトル
    
    
    
    
    ' 不要な線と矢印ブロックの判定および削除
    Dim checkObject As AcadEntity
    Dim checkPoint As Variant
    For Each checkObject In newSelectionSet
        
        If TypeOf checkObject Is AcadLine Then
            
            
            
        ElseIf TypeOf checkObject Is AcadBlockReference Then
            checkPoint = checkObject.InsertionPoint
            If checkPoint = firstPoint Then checkObject.Delete
        End If
        
    Next checkObject
    
    
End Sub

'------------------------------------------------------------------------------
' ## 単位ベクトル算出
'------------------------------------------------------------------------------
Private Sub calculateUnitVector(ByRef point1 As Variant, ByRef point2 As Variant, _
                                ByVal unitvector_x As Double, ByVal unitvector_y As Double)
    
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    
    x1 = point1(0)
    y1 = point1(1)
    x2 = point2(0)
    y2 = point2(1)
    
    unitvector_x = (x2 - x1) / sqrt((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    unitvector_y = (y2 - y1) / sqrt((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
    
End Sub
