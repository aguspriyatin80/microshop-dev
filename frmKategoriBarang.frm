VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KATEGORIBARANG 
   Caption         =   "KATEGORI BARANG"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   OleObjectBlob   =   "frmKategoriBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KATEGORIBARANG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersihkan()
Me.txtNamaKategori.Text = ""
kodeOtomatis
Me.txtNamaKategori.SetFocus
Me.cmdSimpan.Caption = "SIMPAN"
End Sub
Private Sub cboIDKategori_Change()
Dim rg As Range
Dim cIndex As Variant
Dim kode As String
On Error Resume Next
lastRow = Sheets("KATEGORI_BARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set rg = Sheets("KATEGORI_BARANG").Range("A2:B" & lastRow)
kode = Me.txtID.Value
cIndex = Application.Match(kode, rg.Columns(1), 1)
If Len(cIndex) > 0 Then
    Me.txtNamaKategori.Text = rg.Cells(cIndex, 2).Value
End If
End Sub
Sub kodeOtomatis()
Set aktifSheet = Sheets("KATEGORI_BARANG")
lastRow = aktifSheet.Cells(Rows.Count, 2).End(xlUp).Row + 1

kode_otomatis = "CAT" & Format(Right(aktifSheet.Cells(lastRow - 1, 1), 3) + 1, "0##")
Me.txtID.Text = kode_otomatis
End Sub
Private Sub cmdBatal_Click()
bersihkan
Me.cmdHapus.Enabled = False
End Sub
Private Sub cmdHapus_Click()
Dim baris As Long
Dim dbarea As Range, hc As Range
baris = getBaris
Set dbarea = Sheets("KATEGORI_BARANG").Range("A:A" & baris)
Set hc = dbarea.Find(Me.txtID.Text, , xlValues, xlWhole)
If Not hc Is Nothing Then
    hc.EntireRow.Delete
End If
End Sub
Private Sub cmdSimpan_Click()
Dim baris As Integer
If Me.cmdSimpan.Caption = "SIMPAN" Then
    Set aktifSheet = Sheets("KATEGORI_BARANG")
    lastRow = aktifSheet.Cells(Rows.Count, 2).End(xlUp).Row + 1
    If Me.txtNamaKategori.Text <> "" Then
        aktifSheet.Cells(lastRow, 1) = Me.txtID.Text
        aktifSheet.Cells(lastRow, 2) = txtNamaKategori
    Else
        MsgBox "Nama kategori tidak boleh kosong!"
        Me.txtNamaKategori.SetFocus
    End If
Else
    Set aktifSheet2 = Sheets("KATEGORI_BARANG").Range("A:A")
    Set cari = aktifSheet2.Find(Me.txtID.Value, LookAt:=xlWhole)
        If Not cari Is Nothing Then
            baris = cari.Row
            aktifSheet2.Cells(baris, 2) = Me.txtNamaKategori
        End If
End If
tampilkanTabel
bersihkan
Me.cmdHapus.Enabled = False
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.ListBox1.ListIndex = -1 Then Exit Sub
Me.txtID.Value = Me.ListBox1.Column(0)
Me.txtNamaKategori.Value = Me.ListBox1.Column(1)
Me.cmdHapus.Enabled = True
Me.cmdSimpan.Caption = "UPDATE"
End Sub
Private Sub UserForm_Activate()
kodeOtomatis
Me.txtID.Enabled = False
tampilkanTabel
End Sub
Private Sub UserForm_Initialize()
Set aktifSheet = Sheets("KATEGORI_BARANG")
aktifSheet.Select
lastRow = aktifSheet.Cells(Rows.Count, 2).End(xlUp).Row + 1
tampilkanTabel
Me.cmdHapus.Enabled = False
End Sub
Sub tampilkanTabel()
Set aktifSheet = Sheets("KATEGORI_BARANG")
aktifSheet.Select
Me.ListBox1.ColumnCount = 2
Me.ListBox1.ColumnHeads = True
Me.ListBox1.ColumnWidths = "70;90"
lastRow = Sheets("KATEGORI_BARANG").Cells(Rows.Count, 1).End(xlUp).Row
Me.ListBox1.RowSource = "A2:B" & lastRow
End Sub
