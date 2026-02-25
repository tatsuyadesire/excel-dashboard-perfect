Attribute VB_Name = "Módulo1"
Sub AjustarZoomDashboard()
 Dim ws As Worksheet
 Set ws = ThisWorkbook.Sheets("Dashboard") ' Insira a planilha pelo nome da aba
 
 'Selecionar o range do Dashboard e nomear. Pode ser necessário ajudar para se adequar ao resultado final que você quiser. Recomendo fazer testes selecionando um intervalo maior ou menor dependendo do resultado buscado.
 
 ' Selecione o range do dashboard
 ws.Activate
 ws.Range("RangeDash").Select ' altere pelo nome do range
 
 ' Configurações de exibição do Dashboard
 With Application
 .DisplayFullScreen = True
 .DisplayFormulaBar = False
 .DisplayStatusBar = False
 .CommandBars("Worksheet Menu Bar").Enabled = False
 End With
 
 ' Ajuste de zoom e rolagem
 With ActiveWindow
 .Zoom = True
 .ScrollRow = 1
 .ScrollColumn = 1
 End With
 
 ' Tirar o foco da seleção (opcional, apenas para estética)
 ws.Range("A1").Select
End Sub
'Para executar automaticamente ao abrir a planilha

Private Sub Workbook_Open()

 Call AjustarZoomDashboard

End Sub

'Para ajustar automaticamente ao redimensionar a tela/trocar de monitor

Private Sub Workbook_WindowResize(ByVal Wn As Window)

 Call AjustarZoomDashboard

End Sub
