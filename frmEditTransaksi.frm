VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EDIT 
   Caption         =   "EDIT TRANSAKSI"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10920
   OleObjectBlob   =   "frmEditTransaksi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
On Error Resume Next
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set stoknya = Sheets("DATABARANG").Range("b2:b" & lastRow).Find(what:=Me.txtID.Text, LookIn:=xlValues)
Set trx = Sheets("NOTA").Range("b8:b" & lastRow).Find(what:=Me.txtID.Text, LookIn:=xlValues)
Set sementara = Sheets("SEMENTARA").Range("a2:a" & lastRow).Find(what:=Me.txtNama.Text, LookIn:=xlValues)
baris = trx.Row
baris1 = sementara.Row
If MsgBox("Apakah anda yakin akan menghapus transaksi " & Me.txtNama.Text & "?", vbYesNo + vbQuestion, "MICROSHOP") = vbYes Then
    Sheets("NOTA").Cells(baris, "a").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "b").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "c").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "d").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "e").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "f").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "g").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "h").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "i").Delete Shift:=xlUp
    Sheets("NOTA").Cells(baris, "j").Delete Shift:=xlUp
    stoknya.Offset(0, 5).Value = Val(Me.txtStok.Text) + Val(Me.txtJumlah.Text)
    Sheets("SEMENTARA").Cells(baris1, "a").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "b").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "c").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "d").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "e").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "f").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "g").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "h").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "i").Delete Shift:=xlUp
    Sheets("SEMENTARA").Cells(baris1, "j").Delete Shift:=xlUp
    Unload Me
Exit Sub
End If
Me.txtID = ""
Me.txtNama = ""
Me.txtSatuan = ""
Me.txtHarga = ""
Me.txtJumlah = ""
Me.txtJumlahBaru = ""
Me.txtTotal = ""
Me.txtStok = ""
DASHBOARD2.tampil
End Sub
Private Sub cmdUpdate_Click2()
If Val(Me.txtJumlahBaru.Text) > Val(Me.txtStok.Text) Then
MsgBox "Stok tidak mencukupi!"
Me.txtJumlahBaru.SetFocus
Exit Sub
Else
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set stoknya = Sheets("DATABARANG").Range("b2:b" & lastRow).Find(what:=Me.txtID.Text, LookIn:=xlValues)
Set trx = Sheets("NOTA").Range("b8:b" & lastRow).Find(what:=Me.txtID.Text, LookIn:=xlValues)
Set sementara = Sheets("SEMENTARA").Range("a2:a" & lastRow).Find(what:=Me.txtNama.Text, LookIn:=xlValues)
baris = trx.Row
baris1 = sementara.Row
    If Me.txtJumlahBaru.Value <> "" Then
    Sheets("SEMENTARA").Cells(baris1, 2).Value = Me.txtJumlahBaru.Value
    Sheets("SEMENTARA").Cells(baris1, 6).Value = Sheets("SEMENTARA").Cells(baris1, 5).Value * Val(Me.txtJumlahBaru.Text)
    Sheets("SEMENTARA").Cells(baris1, 7).Value = Val(Me.txtHarga.Text) * Val(Me.txtJumlahBaru.Text)
    trx.Offset(0, 3).Value = Me.txtHarga.Value
    trx.Offset(0, 4).Value = Me.txtJumlahBaru.Value
    trx.Offset(0, 5).Value = Val(Me.txtJumlahBaru.Text) * Val(Me.txtHarga.Text)
    Else
    MsgBox "Masukkan jumlah baru", vbOKOnly + vbCritical, "INFORMATION"
    Me.txtJumlahBaru.SetFocus
    Exit Sub
    End If
Unload Me
Me.txtID = ""
Me.txtNama = ""
Me.txtSatuan = ""
Me.txtHarga = ""
Me.txtJumlah = ""
Me.txtJumlahBaru = ""
Me.txtTotal = ""
Me.txtStok = ""
End If
End Sub
Private Sub cmdUpdate_Click()
If Val(Me.txtJumlahBaru.Text) > Val(Me.txtStok.Text) Then
MsgBox "Stok tidak mencukupi!"
Me.txtJumlahBaru.SetFocus
Exit Sub
Else
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set stoknya = Sheets("DATABARANG").Range("b2:b" & lastRow).Find(what:=Me.txtID.Text, LookIn:=xlValues)
Set trx = Sheets("NOTA").Range("b8:b" & lastRow).Find(what:=Me.txtID.Text, LookIn:=xlValues)
Set sementara = Sheets("SEMENTARA").Range("a2:a" & lastRow).Find(what:=Me.txtNama.Text, LookIn:=xlValues)
baris = trx.Row
baris1 = sementara.Row
    If Me.txtJumlahBaru.Value <> "" Then
    Sheets("SEMENTARA").Cells(baris1, 7).Value = Val(Me.txtHarga.Text) * Val(Me.txtJumlahBaru.Text)
    Sheets("SEMENTARA").Cells(baris1, 2).Value = Me.txtJumlahBaru.Value
    Sheets("SEMENTARA").Cells(baris1, 6).Value = Sheets("SEMENTARA").Cells(baris1, 5).Value * Val(Me.txtJumlahBaru.Text)
    trx.Offset(0, 3).Value = Me.txtHarga.Value
    trx.Offset(0, 4).Value = Me.txtJumlahBaru.Value
    trx.Offset(0, 5).Value = Val(Me.txtJumlahBaru.Text) * Val(Me.txtHarga.Text)
    Else
    MsgBox "Masukkan jumlah baru", vbOKOnly + vbCritical, "APLIKASI KASIR"
    Me.txtJumlahBaru.SetFocus
    Exit Sub
    End If
Unload Me
Me.txtID = ""
Me.txtNama = ""
Me.txtSatuan = ""
Me.txtHarga = ""
Me.txtJumlah = ""
Me.txtJumlahBaru = ""
Me.txtTotal = ""
Me.txtStok = ""
End If
End Sub
Private Sub txtHarga_Change()
Me.txtTotal.Value = Val(Me.txtJumlahBaru.Value) * Val(Me.txtHarga.Value)
End Sub
Private Sub txtJumlahBaru_Change()
Me.txtTotal.Value = Val(Me.txtJumlahBaru.Value) * Val(Me.txtHarga.Value)
End Sub
Private Sub UserForm_Activate()
Dim xTextBox As Object
Me.txtJumlahBaru.SetFocus
With txtJumlahBaru
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub UserForm_Terminate()
DASHBOARD2.txtScan.SetFocus
End Sub
