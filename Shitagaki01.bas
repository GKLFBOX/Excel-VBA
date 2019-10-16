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
    secondPoint = ThisDrawing.Utility.GetPoint(fistPoint, "引出線の2点目を指定 [Cancel(ESC)]")
    
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
    
    
End Sub
