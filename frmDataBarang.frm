VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DATABARANG 
   Caption         =   "MANAJEMEN DATA BARANG"
   ClientHeight    =   9075.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13755
   OleObjectBlob   =   "frmDataBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DATABARANG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Worksheet
Dim lokasi_foto As String
Function kodeOtomatis() As String
Set aktifSheet = Sheets("DATABARANG")
akhirBaris = aktifSheet.Cells(Rows.Count, 2).End(xlUp).Row
noUrutTerakhir = aktifSheet.Range("B" & akhirBaris).Value
If Len(noUrutTerakhir) = 5 Then
    noUrut = noUrutTerakhir + 1
Else
Do While Len(noUrutTerakhir) > 5
    akhirBaris = akhirBaris - 1
    noUrutTerakhir = aktifSheet.Range("B" & akhirBaris).Value
    noUrut = noUrutTerakhir + 1
Loop
End If
kodeOtomatis = noUrut
End Function
Sub getImagePath()
Dim pathImage As String
pathImage = Application.GetOpenFilename
Me.txtPath.Text = pathImage
End Sub
Private Sub cmdHapus_Click()
Dim baris As Long
Dim dbarea As Range, hc As Range, aktifSheet As Worksheet
Set aktifSheet = Sheets("DATABARANG")
aktifSheet.Activate
baris = aktifSheet.Cells(Rows.Count, 1).End(xlUp).Row
Set dbarea = aktifSheet.Range("B2:B" & baris)
Set hc = dbarea.Find(Me.txtID.Text, , xlValues, xlWhole)
If MsgBox("Apakah anda yakin akan menghapus barang " & Me.txtNama.Text & "?", vbYesNo + vbQuestion, "MICROSHOP") = vbYes Then
    If Not hc Is Nothing Then
        hc.EntireRow.Delete
    End If
End If
REFRESH
End Sub
Private Sub cmdRemoveImage_Click()
With Application.FileDialog(msoFileDialogFilePicker)
imgBarang.Picture = LoadPicture(vbNullString)
lokasi_foto = ""
End With
Me.txtPath.Text = lokasi_foto
Me.Repaint
End Sub
Private Sub imgBarang_Click()
Dim fd As Boolean
fd = fd Xor True
With Me.imgBarang
    If fd Then
        If .Width < 128 Then
            .Left = .Left - 25
            .Width = .Width + 50
            .Top = .Top - 15
            .Height = .Height + 50
           fd = False
        Else
            .Width = 78
            .Height = 72
            .Left = 420
            .Top = 12
        End If
    End If
End With
End Sub
Private Sub txtCari_Enter()
ph_enter txtCari
End Sub
Private Sub txtCari_Exit(ByVal Cancel As MSForms.ReturnBoolean)
ph_exit txtCari
End Sub
Public Sub placeholder()
txtCari.Tag = "Ketik nama barang atau kode barangnya ..."
End Sub
Public Sub ph_enter(ctrl As MSForms.TextBox)
If ctrl.Value = ctrl.Tag Then
    ctrl.Value = ""
    ctrl.ForeColor = vbBlack
End If
End Sub
Public Sub ph_exit(ctrl As MSForms.TextBox)
If Len(ctrl.Value) = 0 Then
    ctrl.Value = ctrl.Tag
    ctrl.ForeColor = &H80000011
End If
End Sub
Sub CARIBARANG()
On Error Resume Next
If Me.txtCari.Text = "" Or Me.txtCari.Tag = Me.txtCari.Value Then Exit Sub
Set ws = Sheets("DATABARANG")
ws.Activate
With ws
    lRow = .Range("A" & .Rows.Count).End(xlUp).Row
    Set Rng = .Range("A1:O" & lRow)
    .AutoFilterMode = False
    With Rng
        .AutoFilter Field:=2, Criteria1:=Me.txtCari.Value
        Set rngA = .Offset(1, 0).SpecialCells(xlCellTypeVisible)
    End With
    .AutoFilterMode = False
    With Rng
        .AutoFilter Field:=3, Criteria1:="*" & Me.txtCari.Text & "*"
        Set rngB = .Offset(1, 0).SpecialCells(xlCellTypeVisible)
    End With
    .AutoFilterMode = False
    Rng.Offset(1, 0).EntireRow.Hidden = True
    Union(rngA, rngB).EntireRow.Hidden = False
    Sheets("HASILFILTER").Cells.clear
    Sheets("DATABARANG").Range("A1:O" & Cells(Rows.Count, 1).End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy _
    Destination:=Sheets("HASILFILTER").Range("A1")
    lastRow = Sheets("HASILFILTER").Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow - 1 > 2 Then
        Sheets("HASILFILTER").Range("A2").Value = 1
        Sheets("HASILFILTER").Range("A3").Value = 2
        Sheets("HASILFILTER").Range("A2:A3").AutoFill Destination:=Sheets("HASILFILTER").Range("A2:A" & lastRow)
    ElseIf lastRow - 1 = 2 Then
        Sheets("HASILFILTER").Range("A2").Value = 1
        Sheets("HASILFILTER").Range("A3").Value = 2
    ElseIf lastRow - 1 = 1 Then
        Sheets("HASILFILTER").Range("A2").Value = 1
    ElseIf lastRow - 1 = 0 Then
        Sheets("HASILFILTER").Range("A2").Value = ""
        Sheets("HASILFILTER").Range("C2").Value = "Data tidak ditemukan"
    End If
    Rng.Offset(1, 0).EntireRow.Hidden = False
    Sheets("HASILFILTER").Cells.EntireColumn.AutoFit
    Sheets("HASILFILTER").Select
    If Me.txtCari.Text = "" Or Len(Me.txtCari.Text) = 0 Then
        Me.ListBox1.RowSource = "DATABARANG!a2:o" & Cells(Rows.Count, 1).End(xlUp).Row + 1
    Else
        Me.ListBox1.RowSource = "HASILFILTER!a2:o" & Cells(Rows.Count, 1).End(xlUp).Row + 1
    End If
End With
End Sub
Function GetNumeric(CellRef As String)
Dim StringLength As Integer
StringLength = Len(CellRef)
For i = 1 To StringLength
If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
Next i
GetNumeric = result
End Function
Sub myDataTable()
Set ws = Sheets("DATABARANG")
ws.Activate
ws.AutoFilterMode = False
Me.ListBox1.ColumnCount = 15
Me.ListBox1.ColumnHeads = True
Me.ListBox1.ColumnWidths = "30;90;250;50;70;70;40;60;100;0;70;80;60;200;60"
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Me.ListBox1.RowSource = "A2:o" & lastRow
End Sub
Sub kosongkan()
Me.imgBarang.Picture = LoadPicture(ThisWorkbook.Path & "\noimage.jpg")
Me.txtID.Value = ""
Me.txtNama.Value = ""
Me.txtSatuan.Value = ""
Me.txtBeli.Value = ""
Me.txtJual.Value = ""
Me.txtStok.Value = ""
Me.txtStok.Enabled = True
Me.txtMinimalStok.Value = ""
Me.txtLokasi.Value = ""
Me.cboKategori.Value = "UNCATEGORIZED"
Me.txtPath.Text = ""
End Sub
Sub cekstok()
On Error Resume Next
Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim MinStok As Long
Set ws = ThisWorkbook.Sheets("DATABARANG")
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
Me.tbCekStok.clear
For i = 1 To lastRow
If Val(ws.Cells(i, 7).Value) <= Val(ws.Cells(i, 8).Value) Then
    If i = 1 Then
        tbCekStok.AddItem ws.Cells(i, 2).Value & " - " & ws.Cells(i, 3).Value & ""
    Else
        tbCekStok.AddItem ws.Cells(i, 2).Value & " - " & ws.Cells(i, 3).Value & " - Stok Tersisa : " & ws.Cells(i, 7).Value
    End If
End If
Next
End Sub
Private Sub cmdBatal_Click()
kosongkan
Me.txtID.Enabled = True
Me.txtID.SetFocus
Me.cmdUpdate.Enabled = False
Me.cmdInput.Enabled = True
DATABARANG.Repaint
End Sub
Private Sub cmdCari_Click()
If Me.txtCari.Tag = Me.txtCari.Value Then Exit Sub
Set ws = Sheets("DATABARANG")
ws.Activate
ws.AutoFilterMode = False
With ws
    lRow = .Range("A" & .Rows.Count).End(xlUp).Row
    Set Rng = .Range("A1:O" & lRow)
    .AutoFilterMode = False
    With Rng
        On Error Resume Next
        .AutoFilter Field:=2, Criteria1:=Me.txtCari.Text
        Set rngA = .Offset(1, 0).SpecialCells(xlCellTypeVisible)
        Exit Sub
    End With
    .AutoFilterMode = False
    With Rng
        On Error Resume Next
        .AutoFilter Field:=3, Criteria1:="*" & Me.txtCari.Text & "*"
        Set rngB = .Offset(1, 0).SpecialCells(xlCellTypeVisible)
        Exit Sub
    End With
    .AutoFilterMode = False
    Rng.Offset(1, 0).EntireRow.Hidden = True
    Union(rngA, rngB).EntireRow.Hidden = False
    Sheets("HASILFILTER").Cells.clear
    Sheets("DATABARANG").Range("A1:N" & Cells(Rows.Count, 1).End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy _
    Destination:=Sheets("HASILFILTER").Range("A1")
    With Worksheets("DATABARANG")
    lastRow = Sheets("HASILFILTER").Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow - 1 > 2 Then
        Sheets("HASILFILTER").Range("A2").Value = 1
        Sheets("HASILFILTER").Range("A3").Value = 2
        Sheets("HASILFILTER").Range("A2:A3").AutoFill Destination:=Sheets("HASILFILTER").Range("A2:A" & lastRow)
    ElseIf lastRow - 1 = 2 Then
        Sheets("HASILFILTER").Range("A2").Value = 1
        Sheets("HASILFILTER").Range("A3").Value = 2
    ElseIf lastRow - 1 = 1 Then
        Sheets("HASILFILTER").Range("A2").Value = 1
    ElseIf lastRow - 1 = 0 Then
        Sheets("HASILFILTER").Range("A2").Value = ""
        Sheets("HASILFILTER").Range("C2").Value = "Data tidak ditemukan"
    End If
    End With
    Rng.Offset(1, 0).EntireRow.Hidden = False
    Sheets("HASILFILTER").Cells.EntireColumn.AutoFit
    Sheets("HASILFILTER").Select
    If Me.txtCari.Text = "" Or Len(Me.txtCari.Text) = 0 Then
        Me.ListBox1.RowSource = "DATABARANG!a2:o" & Cells(Rows.Count, 1).End(xlUp).Row
    Else
        Me.ListBox1.RowSource = "HASILFILTER!a2:o" & Cells(Rows.Count, 1).End(xlUp).Row
    End If
End With
End Sub
Private Sub cmdImage_Click()
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Filters.Add "Foto", "*.jpg;*.jpeg"
    .Title = "Choose File"
    If .Show = -1 Then
        imgBarang.Picture = LoadPicture(.SelectedItems(1))
        lokasi_foto = .SelectedItems(1)
    End If
End With
Me.txtPath.Text = lokasi_foto
Me.Repaint
End Sub
Private Sub cmdInput_Click()
simpan
End Sub
Sub REFRESH2()
Me.txtCari.Text = ""
Sheets("DATABARANG").Activate
Sheets("DATABARANG").AutoFilterMode = False
Set ws = Sheets("DATABARANG")
lRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
Sheets("DATABARANG").Range("A2:o" & lRow).EntireRow.Hidden = False
Me.ListBox1.RowSource = "DATABARANG!a2:o" & Cells(Rows.Count, 1).End(xlUp).Row
End Sub
Sub REFRESH()
Me.txtCari.Text = ""
ph_exit txtCari
myDataTable
End Sub
Private Sub cmdKeluar_Click()
Unload Me
End Sub
Private Sub cmdRefresh_Click()
REFRESH
End Sub
Private Sub cmdUpdate_Click()
Dim lastRow As Long
Dim datanya As Range
On Error Resume Next
If Me.txtID.Value = "" Then
    MsgBox "Pilih barangnya dulu!"
    Exit Sub
End If
Me.txtID.Enabled = True
Me.cmdInput.Enabled = True
Sheets("DATABARANG").Activate
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set datanya = Sheets("DATABARANG").Range("B1:o" & lastRow).Find(what:=Me.txtID.Value, LookIn:=xlValues, LookAt:=xlWhole)
    datanya.Offset(0, 1).Value = UCase(Me.txtNama.Value)
    datanya.Offset(0, 2).Value = UCase(Me.txtSatuan.Value)
    datanya.Offset(0, 3).Value = Me.txtBeli.Value
    datanya.Offset(0, 3).NumberFormat = "#,##0"
    datanya.Offset(0, 4).Value = Me.txtJual.Value
    datanya.Offset(0, 4).NumberFormat = "#,##0"
    datanya.Offset(0, 5).Value = Me.txtStok.Value
    datanya.Offset(0, 6).Value = UCase(Me.txtMinimalStok.Value)
    datanya.Offset(0, 7).Value = UCase(Me.txtLokasi.Value)
    datanya.Offset(0, 8).Value = "active"
    datanya.Offset(0, 9).Value = UCase(Me.cboKategori.Value)
    With Application.FileDialog(msoFileDialogFilePicker)
            If Me.txtPath = "" Then
                lokasi_foto = ""
            Else
                lokasi_foto = .SelectedItems(1)
            End If
    End With
    If Application.FileDialog(msoFileDialogFilePicker).SelectedItems.Count > 0 Then
        datanya.Offset(0, 12).Value = lokasi_foto
        Me.imgBarang.Picture = LoadPicture(datanya.Offset(0, 12).Value)
    End If
Me.ListBox1.ColumnCount = 15
Me.ListBox1.RowSource = "DATABARANG!a2:o" & Cells(Rows.Count, 1).End(xlUp).Row
Sheets("DATABARANG").AutoFilterMode = False
kosongkan
cekstok
If Me.txtCari.Text <> "" Then
    CARIBARANG
    DATABARANG.Repaint
End If
Me.txtID.SetFocus
Me.cmdUpdate.Enabled = False
Me.cmdInput.Enabled = True
Me.cmdHapus.Enabled = False
Sheets("DATABARANG").AutoFilterMode = False
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.txtID.Value = 0 Or Me.txtID.Text = "ID Barang" Then
    Me.cmdInput.Enabled = True
    Me.cmdUpdate.Enabled = False
    Me.cmdHapus.Enabled = False
    Exit Sub
End If
Me.txtID.Value = Me.ListBox1.Column(1)
Me.txtNama.Value = Me.ListBox1.Column(2)
Me.txtSatuan.Value = Me.ListBox1.Column(3)
Me.txtBeli.Value = Me.ListBox1.Column(4)
Me.txtJual.Value = Me.ListBox1.Column(5)
Me.txtStok.Value = Me.ListBox1.Column(6)
Me.txtMinimalStok.Value = Me.ListBox1.Column(7)
Me.txtLokasi.Value = Me.ListBox1.Column(8)
Me.cboKategori.Value = Me.ListBox1.Column(10)
Me.cmdInput.Enabled = False
Me.cmdUpdate.Enabled = True
Me.cmdHapus.Enabled = True
Me.txtID.Enabled = False
Me.txtStok.Enabled = False
Me.imgBarang.Picture = LoadPicture("" & Me.ListBox1.Column(13))
Me.txtPath = Me.ListBox1.Column(13)
DATABARANG.Repaint
End Sub
Private Sub txtBeli_Change()
Me.txtBeli.Value = Format(Me.txtBeli, "#,##0")
End Sub
Sub simpan()
Dim barangId As String
Dim namaBarang As String
Dim stok As Long
Dim a As Integer
Dim ada As Boolean
On Error Resume Next
Set ws = Worksheets("DATABARANG")
ThisWorkbook.Sheets("DATABARANG").AutoFilterMode = False
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
If Me.txtID.Text <> "" Then
    barangId = Val(Me.txtID.Value)
Else
    barangId = kodeOtomatis()
End If
namaBarang = Me.txtNama.Value
stok = Val(Me.txtStok.Value)
a = 0
lastRow = Worksheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastRow
    If Worksheets("DATABARANG").Cells(i, 2).Value = barangId Then
        a = a + 1
    End If
Next i
If a > 0 Then
        MsgBox ("Kode Barang sudah ada!")
        Me.txtID.Value = ""
        Me.txtID.SetFocus
        Exit Sub
Else
     '=================simpan
     If Me.txtNama.Text = "" Or _
        Me.txtSatuan.Text = "" Or _
        Me.txtBeli.Text = "" Or _
        Me.txtJual.Text = "" Or _
        Me.txtMinimalStok.Text = "" Or _
        Me.txtStok.Text = "" Then
        MsgBox "Lengkapi semua data!"
        Me.txtNama.SetFocus
        Exit Sub
      End If
      If Me.txtID.Value = "" Then
         ws.Range("a" & lastRow + 1).Value = "=Row()-1"
         ws.Range("b" & lastRow + 1).Value = barangId
         ws.Range("b" & lastRow + 1).NumberFormat = "0"
         ws.Range("c" & lastRow + 1).Value = UCase(Me.txtNama.Value)
         ws.Range("d" & lastRow + 1).Value = UCase(Me.txtSatuan.Value)
         ws.Range("e" & lastRow + 1).Value = Me.txtBeli.Value
         ws.Range("e" & lastRow + 1).NumberFormat = "#,##0"
         ws.Range("f" & lastRow + 1).Value = Me.txtJual.Value
         ws.Range("f" & lastRow + 1).NumberFormat = "#,##0"
         ws.Range("g" & lastRow + 1).Value = Me.txtStok.Value
         ws.Range("h" & lastRow + 1).Value = Me.txtMinimalStok.Value
         ws.Range("i" & lastRow + 1).Value = UCase(Me.txtLokasi.Value)
         ws.Range("j" & lastRow + 1).Value = "active"
         ws.Range("k" & lastRow + 1).Value = UCase(Me.cboKategori.Value)
         ws.Range("l" & lastRow + 1).Value = 0
         ws.Range("m" & lastRow + 1).Value = Me.txtStok.Value
         ws.Range("n" & lastRow + 1).Value = lokasi_foto
         ws.Range("o" & lastRow + 1).Value = 0
     Else
         Me.txtNama.SetFocus
         ws.Range("a" & lastRow + 1).Value = "=Row()-1"
         ws.Range("b" & lastRow + 1).Value = Me.txtID.Value
         ws.Range("b" & lastRow + 1).NumberFormat = "0"
         ws.Range("c" & lastRow + 1).Value = UCase(Me.txtNama.Value)
         ws.Range("d" & lastRow + 1).Value = Me.txtSatuan.Value
         ws.Range("e" & lastRow + 1).Value = Me.txtBeli.Value
         ws.Range("e" & lastRow + 1).NumberFormat = "#,##0"
         ws.Range("f" & lastRow + 1).Value = Me.txtJual.Value
         ws.Range("f" & lastRow + 1).NumberFormat = "#,##0"
         ws.Range("g" & lastRow + 1).Value = Me.txtStok.Value
         ws.Range("h" & lastRow + 1).Value = Me.txtMinimalStok.Value
         ws.Range("i" & lastRow + 1).Value = UCase(Me.txtLokasi.Value)
         ws.Range("j" & lastRow + 1).Value = "active"
         ws.Range("k" & lastRow + 1).Value = UCase(Me.cboKategori.Value)
         ws.Range("l" & lastRow + 1).Value = 0
         ws.Range("m" & lastRow + 1).Value = Me.txtStok.Value
         ws.Range("n" & lastRow + 1).Value = lokasi_foto
         ws.Range("o" & lastRow + 1).Value = 0
         End If
End If
     kosongkan
     REFRESH
     cekstok
     Me.txtID.SetFocus
     '================END SIMPAN
End Sub
Private Sub txtCari_AfterUpdate()
CARIBARANG
End Sub
Private Sub txtJual_Change()
Me.txtJual.Value = Format(Me.txtJual, "#,##0")
End Sub
Private Sub UserForm_Activate()
Me.txtID.SetFocus
myDataTable
cekstok
DATABARANG.Repaint
End Sub
Private Sub UserForm_Initialize2()
Set ws = Sheets("DATABARANG")
ws.Activate
myDataTable
cmdUpdate.Enabled = False
cekstok
lastRow = Sheets("KATEGORI_BARANG").Cells(ws.Rows.Count, 2).End(xlUp).Row
Me.cboKategori.clear
For i = 2 To lastRow
   cboKategori.AddItem Sheets("KATEGORI_BARANG").Cells(i, 2).Value
Next i
Me.cboKategori.Value = Sheets("KATEGORI_BARANG").Range("B2").Value
Me.imgBarang.Picture = LoadPicture(ThisWorkbook.Path & "\noimage.jpg")
End Sub
Private Sub UserForm_Initialize()
Set ws = Sheets("DATABARANG")
Call placeholder
ws.Activate
ws.Cells.EntireColumn.Hidden = False
ws.Cells.EntireRow.Hidden = False
myDataTable
cekstok
ws.AutoFilterMode = False
lastRow = Sheets("DATABARANG").Cells(ws.Rows.Count, 2).End(xlUp).Row
ws.Range("A1:o" & lastRow).EntireRow.Hidden = False
cmdUpdate.Enabled = False
lastRow = Sheets("KATEGORI_BARANG").Cells(ws.Rows.Count, 2).End(xlUp).Row
Me.cboKategori.clear
For i = 2 To lastRow
   cboKategori.AddItem Sheets("KATEGORI_BARANG").Cells(i, 2).Value
Next i
Me.cboKategori.Value = Sheets("KATEGORI_BARANG").Range("B2").Value
Me.imgBarang.Picture = LoadPicture(ThisWorkbook.Path & "\noimage.jpg")
End Sub
