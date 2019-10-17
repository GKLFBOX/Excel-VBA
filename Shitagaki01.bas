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
    
    ' 新たに引出線を引く
    Dim firstPoint As Variant
    Dim secondPoint As Variant
    firstPoint = ThisDrawing.Utility.GetPoint(, "引出線の1点目を指定 [Cancel(ESC)]")
    secondPoint = ThisDrawing.Utility.GetPoint(firstPoint, "引出線の2点目を指定 [Cancel(ESC)]")
    
    Dim pickPoint(5) As Double
    Dim i As Long
    For i = 0 To 2
        pickPoint(i) = firstPoint(i)
        pickPoint(i + 3) = secondPoint(i)
    Next i
    
    Dim newLeader As AcadLeader
    Dim leaderType As Integer
    Dim annotationObject As AcadObject
    leaderType = acLineWithArrow
    Set annotationObject = Nothing
    
    Set newLeader = ThisDrawing.ModelSpace.AddLeader(pickPoint, annotationObject, leaderType)
    newLeader.StyleName = "矢印スタイル名"
    
    ' 引出線作図の2点を利用して選択セットを作成
    Dim newSelectionSet As AcadSelectionSet
    Set newSelectionSet = ThisDrawing.SelectionSets.Add("NewSelSet")
    
    Dim selectMode As Integer
    selectMode = acSelectionSetWindow
    newSelectionSet.Select selectMode, firstPoint, secondPoint
    
    ' 不要な線と矢印ブロックの判定および削除
    Dim checkObject As AcadEntity
    For Each checkObject In newSelectionSet
        
        If TypeOf checkObject Is AcadLine Then
            
        End If
        
        If TypeOf checkObject Is AcadBlockReference Then
            If checkObject.InsertionPoint = firstPoint Then checkObject.Delete
        End If
        
    Next checkObject
    
    
End Sub
