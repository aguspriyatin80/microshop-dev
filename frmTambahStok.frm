VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TAMBAHSTOK 
   Caption         =   "TAMBAH STOK"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5370
   OleObjectBlob   =   "frmTambahStok.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TAMBAHSTOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboKodeBarang_Change()
On Error Resume Next
Dim rg As Range
Dim cIndex As Variant
Dim kode As Double
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set rg = Sheets("DATABARANG").Range("B2:i" & lastRow)
kode = Val(Me.cboKodeBarang.Value)
cIndex = Application.Match(kode, rg.Columns(1), 0)
If Len(cIndex) > 0 Then
    Me.txtStokAwal.Text = rg.Cells(cIndex, 6).Value
    Me.txtStokTerjual.Text = rg.Cells(cIndex, 11).Value
    Me.txtSisaStok.Text = rg.Cells(cIndex, 12).Value
    Me.txtStokAkhir.Text = Val(Me.txtSisaStok.Text) + Val(Me.txtStokMasuk.Text)
    Me.txtStokSudahMasuk.Value = rg.Cells(cIndex, 14).Value
    Me.txtStokMasuk.SetFocus
End If
End Sub
Private Sub cboKodeBarang_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim UsedKey As String
    Select Case KeyCode
    Case 8: UsedKey = "RETURN"
    Case 46: UsedKey = "DELETE"
    Case Else
    End Select
    If UsedKey <> "" Then
    KeyCode = 0
    MsgBox "You are not allowed to use the " & UsedKey & "-key in this textbox", 48, "SORRY"
    End If
End Sub
Private Sub cmdCariKode_Click()
CARIKODEBARANG.Show
End Sub
Private Sub cmdUpdateStok_Click()
If Val(Me.txtStokMasuk.Value) < 0 Or Me.txtStokMasuk.Text = "" Then
    MsgBox "Masukkan jumlah stok masuk"
    Me.txtStokMasuk.Text = ""
    Me.txtStokMasuk.SetFocus
    Exit Sub
End If
Me.cboKodeBarang.Enabled = True
Sheets("DATABARANG").Activate
Sheets("DATABARANG").AutoFilterMode = False
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set datanya = Sheets("DATABARANG").Range("a2:m" & lastRow).Find(what:=Me.cboKodeBarang.Value, LookIn:=xlValues)
totalStokMasuk = Val(Me.txtStokSudahMasuk.Text) + Val(Me.txtStokMasuk.Text)
datanya.Offset(0, 13).Value = totalStokMasuk
datanya.Offset(0, 11).Value = Val(Me.txtStokAwal.Text) + totalStokMasuk - Val(Me.txtStokTerjual)
MsgBox "Update stock success!"
Unload Me
End Sub
Private Sub txtStokMasuk_Change()
Me.txtStokAkhir.Value = Val(Me.txtSisaStok.Value) + Val(Me.txtStokMasuk.Value)
End Sub
Private Sub UserForm_Activate()
Me.txtStokAwal.Enabled = False
Me.txtStokAkhir.Enabled = False
Me.txtSisaStok.Enabled = False
Me.txtStokTerjual.Enabled = False
End Sub
Private Sub UserForm_Initialize()
Me.cboKodeBarang.clear
kodeAwal = Sheets("DATABARANG").Range("b2").Value
lastRow = Val(Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row)
For i = 2 To lastRow
Me.cboKodeBarang.AddItem Sheets("DATABARANG").Range("b" & i).Text
Next i
Me.cboKodeBarang.Text = kodeAwal
End Sub
