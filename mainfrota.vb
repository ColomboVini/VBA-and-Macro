Private Sub btnCriar_Click()
ThisWorkbook.Worksheets("Sheets").Activate
Range("A2").Select

Do
    If Not (IsEmpty(ActiveCell)) Then
        ActiveCell.Offset(1, 0).Select
    End If
Loop Until IsEmpty(ActiveCell) = True

ActiveCell.Value = cboMatricula.Value
ActiveCell.Offset(0, 1).Value = txtNome.Value
ActiveCell.Offset(0, 2).Value = cboPlaca.Value
ActiveCell.Offset(0, 3).Value = txtModelo.Value
ActiveCell.Offset(0, 4).Value = txtData.Value
ActiveCell.Offset(0, 5).Value = txtHora.Value
ActiveCell.Offset(0, 6).Value = txtMissao.Value
ActiveCell.Offset(0, 7).Value = txtKM.Value

cboMatricula.Value = Empty
txtData.Value = Empty
txtHora.Value = Empty
txtMissao.Value = Empty
txtKM.Value = Empty
txtModelo = Empty
txtNome.Value = Empty
cboPlaca.Value = Empty

txtNome.SetFocus
End Sub


Private Sub btnPesquisar_Click()
If cboMatricula.Text = "" Then
MsgBox "Digite a Matricula"
txtNome.SetFocus
GoTo Linha1
End If
With Worksheets("Sheets").Range("A:A")
Set c = .Find(cboMatricula.Value, LookIn:=xlValues, LookAt:=xlPart)
If Not c Is Nothing Then
c.Activate
cboMatricula.Value = c.Value
txtNome.Value = c.Offset(0, 1).Value
cboPlaca.Value = c.Offset(0, 2).Value
txtModelo.Value = c.Offset(0, 3).Value
txtData.Value = c.Offset(0, 4).Value
txtHora.Value = c.Offset(0, 5).Value
txtMissao.Value = c.Offset(0, 6).Value
txtKM.Value = c.Offset(0, 7).Value

Else
MsgBox "Cadastro nao existe!"
End If
End With
Linha1:
End Sub

Private Sub btnDeletar_Click()
Dim Resp As Integer
With Worksheets("Sheets").Range("A:A")
Set c = .Find(cboMatricula.Value, LookIn:=xlValues, LookAt:=xlWhole)
If Not c Is Nothing Then
Resp = MsgBox("Tem certeza que deseja excluir o registro?", vbYesNo, "Confirmação")
If Resp = vbYes Then
c.Select
Selection.EntireRow.Delete
cboMatricula.Value = Empty
txtNome.Value = Empty
cboPlaca.Value = Empty
txtModelo.Value = Empty
txtData.Value = Empty
txtHora.Value = Empty
txtMissao.Value = Empty
txtKM.Value = Empty

txtNome.SetFocus
Else
MsgBox "O registro não será excluído!"
End If
Else
MsgBox "Cadastro nao encontrado!"
End If
End With
Exit Sub

End Sub

Private Sub btnSair_Click()
    cboMatricula.Value = Empty
    txtData.Value = Empty
    txtHora.Value = Empty
    txtMissao.Value = Empty
    txtKM.Value = Empty
    txtNome.Value = Empty
    cboPlaca.Value = Empty
    txtModelo.Value = Empty
    Dados.Hide
End Sub


Private Sub btnLimpar_Click()
cboMatricula.Value = Empty
txtData.Value = Empty
txtHora.Value = Empty
txtMissao.Value = Empty
txtKM.Value = Empty
txtNome.Value = Empty
cboPlaca.Value = Empty
txtModelo.Value = Empty
txtNome.SetFocus

End Sub

Private Sub cboMatricula_Change()
Sheets("Sheets").Select

If cboMatricula.Value = "" Then
    txtNome.Value = Empty
    ElseIf cboMatricula.Value = "<<value>>" Then
    txtNome.Value = Range("<<value>>").Value
End If
Sheets("Sheets").Select
End Sub

Private Sub cboPlaca_Change()
Sheets("Sheets").Select

If cboPlaca.Value = "" Then
    txtModelo.Value = Empty

End Sub
