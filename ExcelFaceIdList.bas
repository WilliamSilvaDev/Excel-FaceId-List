Attribute VB_Name = "ExcelFaceIdList"
' ***************************************************************
' Dê os créditos à @WillamSilvaDev por disponibilizar esse guia
' https://williamsilvadev.github.io
' https://github.com/WilliamSilvaDev/Excel-FaceId-List
' ***************************************************************

Option Explicit

Const APP_NAME = "FaceIDs (Browser)"

' The number of icons to be displayed in a set.
Const ICON_SET = 30

Sub BarOpen()

    MsgBox "Este procedimento pode demorar alguns minutos, tenha paciência", vbOKOnly, "Aviso"
    
    Dim xBar As CommandBar
    Dim xBarPop As CommandBarPopup
    Dim bCreatedNew As Boolean
    Dim n As Integer, m As Integer
    Dim k As Integer
    
    On Error Resume Next
    ' Try to get a reference to the 'FaceID Browser' toolbar if it exists and delete it:
    Set xBar = CommandBars(APP_NAME)
    On Error GoTo 0
    If Not xBar Is Nothing Then
      xBar.Delete
      Set xBar = Nothing
    End If
    
    Set xBar = CommandBars.Add(Name:=APP_NAME, Temporary:=True) ', Position:=msoBarLeft
    With xBar
      .Visible = True
      '.Width = 80
      For k = 0 To 4 ' 5 dropdowns, each for about 1000 FaceIDs
        Set xBarPop = .Controls.Add(Type:=msoControlPopup) ', Before:=1
        With xBarPop
          .BeginGroup = True
          If k = 0 Then
            .Caption = "Face IDs " & 1 + 1000 * k & " ... "
          Else
            .Caption = 1 + 1000 * k & " ... "
          End If
          n = 1
          Do
            With .Controls.Add(Type:=msoControlPopup) '34 items * 30 items = 1020 faceIDs
              .Caption = 1000 * k + n & " ... " & 1000 * k + n + ICON_SET - 1
              For m = 0 To ICON_SET - 1
                With .Controls.Add(Type:=msoControlButton) '
                  .Caption = "ID=" & 1000 * k + n + m
                  .FaceId = 1000 * k + n + m
                End With
              Next m
            End With
            n = n + ICON_SET
          Loop While n < 1000 ' or 1020, some overlapp
        End With
      Next k
    End With 'xBar
End Sub
