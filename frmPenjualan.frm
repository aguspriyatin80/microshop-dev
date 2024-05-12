VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DASHBOARD2 
   Caption         =   "TRIPUTRA VER. 1.0"
   ClientHeight    =   9075.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15210
   OleObjectBlob   =   "frmPenjualan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DASHBOARD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lokasi_toko As String
Dim Berhenti As Boolean
Sub selectLastListbox()
Dim i As Long
    For i = 0 To Me.tbTransaksi.ListCount - 1
        Me.tbTransaksi.ListIndex = i - 1
    Next i
End Sub
Sub cboNoTransaksiData()
Dim rango, celda As Range
Set ws = Sheets("REKAP")
ws.AutoFilterMode = False
lastRow = Sheets("REKAP").Cells(Rows.Count, 1).End(xlUp).Row
Set rango = Worksheets("REKAP").Range("H2:H" & lastRow)
Me.cboNoTransaksi.clear
For Each celda In rango
    Me.cboNoTransaksi.AddItem celda.Value
Next celda
For i = 0 To Me.cboNoTransaksi.ListCount - 2
    For j = Me.cboNoTransaksi.ListCount - 1 To i + 1 Step -1
        If Me.cboNoTransaksi.List(i) = Me.cboNoTransaksi.List(j) Then 'repeated
            Me.cboNoTransaksi.RemoveItem (j)
        End If
    Next j
Next i
End Sub
Sub AutofitRows()
    Dim c As Range
    Set ws = Sheets("NOTA")
    ws.Activate
    akhir = Sheets("NOTA").Cells(Rows.Count, 1).End(xlUp).Row
    ActiveSheetUsedRange = Range("c8:c" & akhir)
    For Each c In ActiveSheetUsedRange
        If c.WrapText Then c.Rows.AutoFit
    Next c
End Sub
Sub cetakPenjualan()
lastRow = Sheets("REKAP").Cells(Rows.Count, 1).End(xlUp).Row

'Range("G" & lastRow + 1).Formula = "=SUMIF(I2:I & lastRow); " = " & TODAY(); G2:G & lastRow)"
' Declaring variables
    Dim sumRange As Range
    Dim criteriaRange As Range
    Dim criteria As String
    Dim result As Double
' Assigning values to variables
    Set sumRange = Sheets("REKAP").Range("G2:G" & lastRow)
    Set criteriaRange = Sheets("REKAP").Range("i2:i" & lastRow)
    criteria = Format(Sheets("NOURUT").Range("A1"), "dd/mm/yyyy")
' Calculating the result using the SUMIFS function
    result = WorksheetFunction.SumIfs(sumRange, criteriaRange, criteria)
    lblPenjualanHariIni.Caption = Format(result, "#,##0")
End Sub
Sub createNoNota()
Dim noUrut As String
lastRow = Sheets("REKAP").Cells(Rows.Count, 1).End(xlUp).Row
noUrut = Val(Application.WorksheetFunction.Max(Sheets("REKAP").Range("h2:h" & lastRow)) + 1)
TGL = Application.WorksheetFunction.Text(Date, "yymmdd")
tglSkrg = Right(TGL, 2)
If (noUrut = 1) Then
    noUrut = TGL & "001"
Else
    If Mid(noUrut, 5, 2) <> tglSkrg Then
        noUrut = TGL & "001"
    End If
End If
Me.lblNoTransaksi.Caption = noUrut
Me.cboNoTransaksi.clear
cboNoTransaksiData
Me.cboNoTransaksi.AddItem noUrut
Me.cboNoTransaksi.Value = noUrut
End Sub
Sub stokKeluar()
On Error Resume Next
    Dim i As Integer
    For i = 0 To tbTransaksi.ListCount - 1
        'Ambil nilai kode barang, nama barang, dan jumlah dari ListBox "tTransaksi"
        Dim kodeBarang As String
        kodeBarang = tbTransaksi.List(i, 1)
        Dim namaBarang As String
        namaBarang = tbTransaksi.List(i, 2)
        Dim jumlah As Integer
        jumlah = tbTransaksi.List(i, 5)
        'Cari baris pertama yang mengandung nilai nama barang yang diinputkan
        Dim firstRow As Integer
        firstRow = Sheets("DATABARANG").Range("b:b").Find(kodeBarang, LookIn:=xlValues, LookAt:=xlWhole).Row
        'Cari baris terakhir yang mengandung nilai nama barang yang diinputkan
        Dim lastRow As Integer
        'lastRow = Sheets("DATABARANG").Range("b:b").Find(kodeBarang, LookIn:=xlValues, SearchDirection:=xlPrevious).Row
        lastRow = Sheets("DATABARANG").Range("b:b").Find(kodeBarang, LookIn:=xlValues, LookAt:=xlWhole).Row
        'Ambil nilai stok keluar dan sisa stok dari sel aktif
        Dim stokKeluar As Integer
        stokKeluar = Sheets("DATABARANG").Range("L" & firstRow).Value
        Dim jumlahStok As Integer
        stokAwal = Sheets("DATABARANG").Range("G" & firstRow).Value
        stokSudahMasuk = Sheets("DATABARANG").Range("O" & firstRow).Value
        'Hitung nilai stok keluar dan sisa stok yang baru
        Dim newStokKeluar As Integer
        newStokKeluar = stokKeluar + jumlah
        Dim newSisaStok As Integer
        newSisaStok = stokAwal + stokSudahMasuk - newStokKeluar
        'Update nilai stok keluar dan sisa stok di lembar kerja
        Sheets("DATABARANG").Range("L" & firstRow - 1).Value = newStokKeluar
        Sheets("DATABARANG").Range("M" & lastRow - 1).Value = newSisaStok
    Next i
End Sub
Sub tampil()
On Error Resume Next
Set ws = Sheets("NOTA")
ws.Activate
akhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 0).Row
Me.tbTransaksi.ColumnHeads = True
Me.tbTransaksi.ColumnWidths = "30 pt;80 pt;220 pt;0 pt;60 pt;30 pt;60 pt"
Me.tbTransaksi.clear
Me.tbTransaksi.RowSource = "a8:g" & akhir + 1
End Sub
Sub kosongkan()
Set ws = Sheets("NOTA")
ws.Activate
Sheets("NOTA").Range("a8:j100").ClearContents
Me.txtBayar.Value = ""
Me.txtKembali.Value = ""
Me.txtScan.SetFocus
Me.lblTotalBelanja.Caption = 0
Me.txtCustomer.Text = ""
Me.tbTransaksi.RowSource = "a8:g8"
End Sub
Sub selesai()
akhir = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
akhir1 = Sheets("REKAP").Cells(Rows.Count, 1).End(xlUp).Row
akhir2 = Sheets("NOTA").Cells(Rows.Count, 1).End(xlUp).Row
akhir3 = Sheets("SEMENTARA").Cells(Rows.Count, 1).End(xlUp).Row
If Sheets("NOTA").Range("c8").Value = "" Then
    MsgBox "Tidak ada transaksi untuk diselesaikan", vbOKOnly + vbCritical, "INFO"
    Exit Sub
ElseIf Me.txtBayar.Value = "" Then
    MsgBox "Jumlah pembayaran masih kosong", vbOKOnly + vbCritical, "APLIKASI KASIR"
    Exit Sub
ElseIf Sheets("NOTA").Range("b" & akhir2 + 3).Value <> "" Then
    MsgBox "Transaksi ini sudah ditutup, silahkan buat transaksi baru atau batalkan dulu!", vbOKOnly + vbCritical, "APLIKASI KASIR"
    Exit Sub
Else
    If Me.txtCustomer.Text = "" Then
        Sheets("NOTA").Range("c6").Value = "UMUM"
    Else
        Sheets("NOTA").Range("c6").Value = Me.txtCustomer.Text
    End If

    Sheets("SEMENTARA").Range("a2:k" & akhir3).Copy
    Sheets("REKAP").Range("a" & akhir1 + 1).PasteSpecial xlPasteValues
    
    Sheets("REKAP").Range("k2:k" & akhir1 + 1).NumberFormat = "0"
    Application.CutCopyMode = False
    Sheets("NOTA").Range("a8:i100").Borders.LineStyle = xlNone
    
    Sheets("NOTA").Range("a" & akhir2, "i" & akhir2).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
    Sheets("NOTA").Range("c4").Value = lblNoTransaksi.Caption
    Sheets("NOTA").Range("g5").Value = JAM.Caption
    Sheets("NOTA").Range("g6").Font.Color = RGB(255, 255, 255)
    Sheets("NOTA").Range("e" & akhir2 + 2).Value = "TOTAL"
    Sheets("NOTA").Range("f" & akhir2 + 2).Value = "Rp"
    Sheets("NOTA").Range("b" & akhir2 + 2).Value = "TOTAL"
    Sheets("NOTA").Range("e" & akhir2 + 3).Value = "BAYAR"
    Sheets("NOTA").Range("f" & akhir2 + 3).Value = "Rp"
    Sheets("NOTA").Range("b" & akhir2 + 3).Value = "BAYAR"
    Sheets("NOTA").Range("e" & akhir2 + 4).Value = "KEMBALI"
    Sheets("NOTA").Range("f" & akhir2 + 4).Value = "Rp"
    Sheets("NOTA").Range("b" & akhir2 + 4).Value = "KEMBALI"
    Sheets("NOTA").Range("g" & akhir2 + 2).Value = DASHBOARD2.txtTotalBelanja.Text
    Sheets("NOTA").Range("g" & akhir2 + 3).Value = DASHBOARD2.txtBayar2.Text
    Sheets("NOTA").Range("g" & akhir2 + 4).Value = DASHBOARD2.txtKembali.Text
    
    Sheets("NOTA").Range("c" & akhir2 + 5).Value = "Terima Kasih."

    Sheets("NOTA").Range("e4").Value = Me.TGL.Caption
    Sheets("NOTA").Range("c8:c" & akhir2).WrapText = True
    Sheets("SEMENTARA").Range("a2:k" & akhir3).ClearContents
    datane
    SaveActiveSheetsAsPDF
    Application.Dialogs(xlDialogPrint).Show
    stokKeluar
    kosongkan
    cetakPenjualan
    createNoNota
    End If
Sheets("DATABARANG").AutoFilterMode = False
ThisWorkbook.Save
End Sub
Private Sub cboNoTransaksi_Change()
Dim sht As Worksheet
Dim shtName As String
If Me.cboNoTransaksi.Value = Me.lblNoTransaksi.Caption Then
    Me.cmdProses.Enabled = True
    Me.cmdPending.Enabled = True
    Me.cmdBatal.Enabled = True
    Set ws = Sheets("NOTA")
    ws.Activate
    akhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    Me.tbTransaksi.ColumnWidths = "30 pt;80 pt;220 pt;0 pt;60 pt;30 pt;60 pt"
    Me.tbTransaksi.ColumnHeads = True
    Me.tbTransaksi.RowSource = "a8:g" & akhir + 1
    totalbelanja
Else
    Me.cmdProses.Enabled = False
    Me.cmdPending.Enabled = False
    Me.cmdBatal.Enabled = False

Me.tbTransaksi.RowSource = ""
For Each sht In ThisWorkbook.Worksheets
    If sht.Name = "REKAP2" Then
        Application.DisplayAlerts = False
        sht.Delete
        Application.DisplayAlerts = True
    End If
Next sht
Set sh = Sheets("REKAP")
sh.AutoFilterMode = False
ish = sh.Range("a" & Application.Rows.Count).End(xlUp).Row
lastRow = sh.Cells(Rows.Count, 1).End(xlUp).Row
Set findData = sh.Range("a1:O" & lastRow).Find(what:=Me.cboNoTransaksi.Value, LookIn:=xlValues, LookAt:=xlWhole)
sh.Range("a1:O" & ish).AutoFilter Field:=8, Criteria1:=Me.cboNoTransaksi.Text
Sheets.Add.Name = "REKAP2"
Sheets("REKAP2").Range("A1").Value = "NO" ' NO.
sh.Range("K1:K" & ish).Copy Sheets("REKAP2").Range("B1") ' KODE BARANG
sh.Range("A1:A" & ish).Copy Sheets("REKAP2").Range("C1") ' NAMA BARANG
sh.Range("C1:C" & ish).Copy Sheets("REKAP2").Range("D1") ' HARGA BELI
sh.Range("D1:D" & ish).Copy Sheets("REKAP2").Range("E1") ' HARGA JUAL
sh.Range("B1:B" & ish).Copy Sheets("REKAP2").Range("F1") ' QTY
sh.Range("G1:G" & ish).Copy Sheets("REKAP2").Range("G1") ' TOTAL
Sheets("REKAP2").Range("E1").Value = " HARGA"
Sheets("REKAP2").Range("G1").Value = "   TOTAL"

Set sh2 = Sheets("REKAP2")
lastRow2 = sh2.Cells(Rows.Count, 2).End(xlUp).Row
sh2.Range("E2:E" & lastRow2).NumberFormat = "#,##0"
sh2.Range("G2:G" & lastRow2).NumberFormat = "#,##0"
    If lastRow2 - 1 > 2 Then
        Sheets("REKAP2").Range("A2").Value = 1
        Sheets("REKAP2").Range("A3").Value = 2
        Sheets("REKAP2").Range("A2:A3").AutoFill Destination:=Sheets("REKAP2").Range("A2:A" & lastRow2)
    ElseIf lastRow2 - 1 = 2 Then
        Sheets("REKAP2").Range("A2").Value = 1
        Sheets("REKAP2").Range("A3").Value = 2
    ElseIf lastRow2 - 1 = 1 Then
        Sheets("REKAP2").Range("A2").Value = 1
    ElseIf lastRow2 - 1 = 0 Then
        Sheets("REKAP2").Range("A2").Value = ""
        Sheets("REKAP2").Range("C2").Value = "Data tidak ditemukan"
    End If
Me.tbTransaksi.ColumnCount = 7
Me.tbTransaksi.ColumnHeads = True
tbTransaksi.ColumnWidths = ("30;80;220;0;60;30;60")
Me.tbTransaksi.RowSource = "REKAP2!A2:g" & lastRow2
totalbelanja2
End If
End Sub
Private Sub cmdBatal_Click()
On Error Resume Next
kosongkan
akhir = Sheets("SEMENTARA").Cells(Rows.Count, 1).End(xlUp).Row
Sheets("SEMENTARA").Range("a2:K" & akhir).ClearContents
Call totalbelanja
Me.txtScan.SetFocus
End Sub
Sub FillCellFromAbove()
For Each cell In selection
    If cell.Value = "" Then
        cell.Value = cell.Offset(-1, 0).Value
    End If
Next cell
End Sub
Private Sub cmdPending_Click()
Dim pendingId As Integer
akhir = Sheets("NOTA").Cells(Rows.Count, 1).End(xlUp).Row
akhir2 = Sheets("PENDING_DETAIL").Cells(Rows.Count, 1).End(xlUp).Row
Sheets("NOTA").Range("a8:J" & akhir + 1).Copy Destination:=Sheets("PENDING_DETAIL").Range("a" & akhir2 + 1)
kosongkan
End Sub
Private Sub cmdProses_Click()
If Sheets("NOTA").Range("A8").Value = "" Then
MsgBox "Belum Ada Transaksi!"
Exit Sub
Else
PEMBAYARAN.Show
End If
End Sub
Sub SaveActiveSheetsAsPDF()
Dim saveLocation As String
no_nota = Sheets("NOTA").Range("c4").Value
If Me.txtCustomer.Text = "" Then
    saveLocation = ThisWorkbook.Path & "\Nota\" & no_nota & "-UMUM.pdf"
Else
    saveLocation = ThisWorkbook.Path & "\Nota\" & no_nota & "-" & Me.txtCustomer.Text & ".pdf"
End If
'Save Active Sheet(s) as PDF
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
    filename:=saveLocation
End Sub
Private Sub CommandButton1_Click()
CARICUSTOMER.Show
End Sub
Private Sub CommandButton2_Click()
DASHBOARD2.Hide
ThisWorkbook.Save
If Windows.Count > 1 Then
    Windows(ThisWorkbook.Name).Activate
    Windows(ThisWorkbook.Name).Application.Visible = True
Else
    Application.Visible = True
End If
Sheets("DATABARANG").AutoFilterMode = False
Sheets("CETAKBARCODE2").Select
CommandBars("Exit Design Mode").Controls(1).Execute
End Sub
Private Sub Label33_Click()
TAMBAHSTOK.Show
End Sub
Private Sub Label34_Click()
SETTINGBARCODE.Show
End Sub
Private Sub Label36_Click()
KATEGORIBARANG.Show
End Sub
Private Sub Label37_Click()
DATACUSTOMER.Show
End Sub
Private Sub Label38_Click()
DASHBOARD2.Hide
ThisWorkbook.Save
Application.Visible = True
Sheets("DATABARANG").AutoFilterMode = False
Sheets("CETAKBARCODE2").Select
CommandBars("Exit Design Mode").Controls(1).Execute
End Sub
Private Sub Label55_Click()
PROFILTOKO.Show
End Sub
Private Sub Label23_Click()
DASHBOARD2.Hide
ThisWorkbook.Save
Application.Visible = True
Sheets("DATABARANG").AutoFilterMode = False
Sheets("MENU").Select
CommandBars("Exit Design Mode").Controls(1).Execute
End Sub
Private Sub Label8_Click()
CARIBARANG.Show
End Sub
Private Sub lblPenjualan_Click()
REKAP.Show
End Sub
Private Sub lblLaporan_Click()
LAPORAN.Show
End Sub
Private Sub lblStokT_Click()
DATABARANG.Show
End Sub
Private Sub lblSettingBarcode_Click()
DASHBOARD2.Hide
ThisWorkbook.Save
If Windows.Count > 1 Then
    Windows(ThisWorkbook.Name).Activate
    Windows(ThisWorkbook.Name).Application.Visible = True
Else
    Application.Visible = True
End If
Sheets("DATABARANG").AutoFilterMode = False
Sheets("CETAKBARCODE2").Select
CommandBars("Exit Design Mode").Controls(1).Execute
End Sub
Private Sub lblStok_Click()
DATABARANG.Show
End Sub
Private Sub RefEdit1_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)
CARIBARANG.Show
End Sub
Private Sub logoToko_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
DASHBOARD2.Hide
ThisWorkbook.Save
If Windows.Count > 1 Then
    Windows(ThisWorkbook.Name).Activate
    Windows(ThisWorkbook.Name).Application.Visible = True
Else
    Application.Visible = True
End If
Sheets("DATABARANG").AutoFilterMode = False
Sheets("CETAKBARCODE2").Select
CommandBars("Exit Design Mode").Controls(1).Execute
End Sub
Private Sub tbTransaksi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
If Me.tbTransaksi.Column(1) = "" Or Me.cboNoTransaksi.Value <> Me.lblNoTransaksi.Caption Then
    Exit Sub
Else
    lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
    Set stok = Sheets("DATABARANG").Range("b2:b" & lastRow).Find(what:=Me.tbTransaksi.Column(1), LookIn:=xlValues)
    EDIT.txtID.Value = Me.tbTransaksi.Column(1)
    EDIT.txtNama.Value = Me.tbTransaksi.Column(2)
    EDIT.txtSatuan.Value = Me.tbTransaksi.Column(3)
    EDIT.txtHarga.Value = Me.tbTransaksi.Column(4)
    EDIT.txtJumlah.Value = Me.tbTransaksi.Column(5)
    EDIT.txtJumlahBaru.Value = Me.tbTransaksi.Column(5)
    EDIT.txtTotal.Value = Me.tbTransaksi.Column(6)
    EDIT.txtStok.Value = stok.Offset(0, 11).Value
    EDIT.txtID.Enabled = False
    EDIT.txtNama.Enabled = False
    EDIT.txtSatuan.Enabled = False
    EDIT.txtHarga.Enabled = True
    EDIT.txtJumlah.Enabled = False
    EDIT.txtTotal.Enabled = False
    EDIT.txtStok.Enabled = False
    EDIT.Show
End If
End Sub
Private Sub tbTransaksi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Then
    On Error Resume Next
    If Me.tbTransaksi.Column(1) = "" Then
        Exit Sub
    Else
        lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
        Set stok = Sheets("DATABARANG").Range("b2:b" & lastRow).Find(what:=Me.tbTransaksi.Column(1), LookIn:=xlValues)
        EDIT.txtID.Value = Me.tbTransaksi.Column(1)
        EDIT.txtNama.Value = Me.tbTransaksi.Column(2)
        EDIT.txtSatuan.Value = Me.tbTransaksi.Column(3)
        EDIT.txtHarga.Value = Me.tbTransaksi.Column(4)
        EDIT.txtJumlah.Value = Me.tbTransaksi.Column(5)
        EDIT.txtJumlahBaru.Value = Me.tbTransaksi.Column(5)
        EDIT.txtTotal.Value = Me.tbTransaksi.Column(6)
        EDIT.txtStok.Value = stok.Offset(0, 11).Value
        EDIT.txtID.Enabled = False
        EDIT.txtNama.Enabled = False
        EDIT.txtSatuan.Enabled = False
        EDIT.txtHarga.Enabled = True
        EDIT.txtJumlah.Enabled = False
        EDIT.txtTotal.Enabled = False
        EDIT.txtStok.Enabled = False
        EDIT.Show
    End If
End If
End Sub
Private Sub txtBayar_Change()
Me.txtBayar.Text = Format(txtBayar.Text, "#,##0")
Me.txtBayar2.Text = Replace(txtBayar.Text, ",", "")
Me.txtBayar2.Text = Replace(txtBayar.Text, ".", "")
txtKembali.Value = Val(txtBayar2.Text) - Val(txtTotalBelanja.Text)
Me.txtKembali.Text = Format(Me.txtKembali.Value, "#,##0")
lblBayar.Caption = Format(txtBayar2.Text, "#,##0")
lblKembali.Caption = Me.txtKembali.Text
End Sub
Private Sub txtBayar_AfterUpdate()
Me.txtBayar2.Text = Replace(txtBayar.Text, ",", "")
Me.txtBayar2.Text = Replace(txtBayar.Text, ".", "")
Me.txtKembali.Value = Val(txtBayar2.Text) - Val(txtTotalBelanja.Text)
Me.txtKembali.Text = Format(Me.txtKembali.Value, "#,##0")
End Sub
Private Sub txtBayar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyEnter Then
    cmdSelesai.SetFocus
End If
End Sub
Private Sub txtCustomer_AfterUpdate()
Me.txtCustomer.Text = UCase(Me.txtCustomer.Text)
End Sub
Private Sub txtScan_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        If Me.txtScan.Text = "" Then
            KeyCode = 0
            Me.txtScan.Text = ""
            Me.txtScan.SetFocus
            'Exit Sub
        Else
            Call isi    'kode lain yang akan di eksekusi
            KeyCode = 0
            Me.txtScan.Text = ""
            Me.txtScan.SetFocus
        End If
    ElseIf KeyCode = vbKeyTab Then
        Me.txtBayar.SetFocus
    ElseIf KeyCode = vbKeyShift Then
        CARIBARANG.Show
    End If
tampil
totalbelanja
'selectLastListbox
End Sub
Private Sub UserForm_Activate()
cek
totalbelanja
jamDigital
cetakPenjualan
Me.tbTransaksi.SetFocus
End Sub
Sub isi()
Sheets("DATABARANG").AutoFilterMode = False
jumlah = 1
akhir = Sheets("NOTA").Cells(Rows.Count, 1).End(xlUp).Row
akhir1 = Sheets("SEMENTARA").Cells(Rows.Count, 1).End(xlUp).Row
akhir2 = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set datanya = Sheets("DATABARANG").Range("b2:b" & akhir2).Find(what:=Me.txtScan.Value, LookIn:=xlValues, LookAt:=xlWhole)
If datanya Is Nothing Then
MsgBox "ID barang tidak ditemukan"
Exit Sub
ElseIf jumlah = "" Then
MsgBox "Jumlah tidak boleh kosong"
Exit Sub
ElseIf Sheets("NOTA").Range("b" & akhir + 2) <> "" Then
MsgBox "Transaksi sudah ditutup,silahkan buat transaksi baru", vbOKOnly + vbCritical, "APLIKASI KASIR"
Exit Sub
Else
    If datanya.Offset(0, 11).Value < 1 Then
        MsgBox "Stok barang kosong!"
        Exit Sub
    End If
    Set cekKode = Sheets("NOTA").Range("b8:b" & akhir).Find(what:=Me.txtScan.Text, LookIn:=xlValues, LookAt:=xlWhole)
    If cekKode Is Nothing Then
        Sheets("NOTA").Range("a" & akhir + 1).Value = "=Row()-7"
        Sheets("NOTA").Range("f" & akhir + 1).Value = jumlah
        Sheets("NOTA").Range("b" & akhir + 1).Value = datanya.Offset(0, 0).Value
        Sheets("NOTA").Range("c" & akhir + 1).Value = datanya.Offset(0, 1).Value
        Sheets("NOTA").Range("d" & akhir + 1).Value = datanya.Offset(0, 2).Value
        Sheets("NOTA").Range("e" & akhir + 1).Value = datanya.Offset(0, 4).Value
        Sheets("NOTA").Range("g" & akhir + 1).Value = datanya.Offset(0, 4).Value * jumlah
        Sheets("NOTA").Range("h" & akhir + 1).Value = Me.lblNoTransaksi.Caption
        Sheets("NOTA").Range("i" & akhir + 1).Value = Format(Date, "dd/mm/yyyy")
        If Me.txtCustomer.Text = "" Then
            Sheets("NOTA").Range("j" & akhir + 1).Value = "UMUM"
        Else
            Sheets("NOTA").Range("j" & akhir + 1).Value = Me.txtCustomer.Value
        End If
        Sheets("SEMENTARA").Range("a" & akhir1 + 1).Value = datanya.Offset(0, 1).Value
        Sheets("SEMENTARA").Range("b" & akhir1 + 1).Value = jumlah
        Sheets("SEMENTARA").Range("c" & akhir1 + 1).Value = datanya.Offset(0, 3).Value
        Sheets("SEMENTARA").Range("d" & akhir1 + 1).Value = datanya.Offset(0, 4).Value
        selisih = datanya.Offset(0, 4).Value - datanya.Offset(0, 3).Value
        Sheets("SEMENTARA").Range("e" & akhir1 + 1).Value = selisih
        Sheets("SEMENTARA").Range("f" & akhir1 + 1).Value = selisih * jumlah
        Sheets("SEMENTARA").Range("G" & akhir1 + 1).Value = datanya.Offset(0, 4).Value * jumlah
        Sheets("SEMENTARA").Range("H" & akhir1 + 1).Value = lblNoTransaksi.Caption
        Sheets("SEMENTARA").Range("i" & akhir1 + 1).Value = Format(Date, "dd/mm/yyyy")
            If Me.txtCustomer.Text = "" Then
                Sheets("SEMENTARA").Range("j" & akhir1 + 1).Value = "UMUM"
            Else
                Sheets("SEMENTARA").Range("j" & akhir1 + 1).Value = Me.txtCustomer.Value
            End If
    Sheets("SEMENTARA").Range("k" & akhir1 + 1).Value = datanya.Offset(0, 0).Value
    Else
    cekKode.Offset(0, 4).Value = cekKode.Offset(0, 4).Value + 1
    cekKode.Offset(0, 5).Value = cekKode.Offset(0, 4).Value * cekKode.Offset(0, 3).Value
   End If
   AutofitRows
End If
Me.txtScan.Value = ""
End Sub
Sub totalbelanja()
On Error Resume Next
Me.txtTotalBelanja.Text = WorksheetFunction.Sum(Sheets("NOTA").Range("g8:g" & Cells(Rows.Count, 1).End(xlUp).Row))
Me.lblTotalBelanja.Caption = Format(Me.txtTotalBelanja.Text, "#,##0")
End Sub
Sub totalbelanja2()
On Error Resume Next
Me.txtTotalBelanja.Text = WorksheetFunction.Sum(Sheets("REKAP2").Range("g2:g" & Cells(Rows.Count, 1).End(xlUp).Row))
Me.lblTotalBelanja.Caption = Format(Me.txtTotalBelanja.Text, "#,##0")
End Sub
Private Sub UserForm_Initialize()
Set ws = Sheets("REKAP")
ws.AutoFilterMode = False
Me.lblNamaToko.Caption = Sheets("PROFIL_TOKO").Range("C1").Value
Me.lblAlamat1.Caption = Sheets("PROFIL_TOKO").Range("C2").Value
Me.lblAlamat2.Caption = Sheets("PROFIL_TOKO").Range("C3").Value
'Me.logoToko.Picture = LoadPicture(lokasi_foto)

With Application.FileDialog(msoFileDialogFilePicker)
lokasi_foto = Sheets("PROFIL_TOKO").Range("C4").Value
End With

'Me.imgLogo.Picture = LoadPicture(ThisWorkbook.Path & "\img_logo.jpg")
'Me.logoToko.Picture = LoadPicture(Sheets("PROFIL_TOKO").Range("C4").Value)

Sheets("DATABARANG").AutoFilterMode = False
Sheets("DATA_CUSTOMER").AutoFilterMode = False

txtScan.Enabled = True
Me.txtScan.SetFocus
Me.lblLogin.Visible = False
Call datane
Sheet6.Range("e2").Value = Date
Me.txtTotalBelanja.Visible = True
Me.txtBayar2.Visible = False
Me.lblBayar.Visible = False
Me.lblKembali.Visible = False
Me.lblNoNota.Visible = False
Me.Label19.Visible = False
Me.Label20.Visible = False
Me.Label22.Visible = False
Me.lblTerjual.Visible = False
Me.lblPendapatan.Visible = False
Me.Label14.Visible = False
Me.Label22.Visible = False
Me.txtBayar.Visible = False
Me.txtKembali.Visible = False
Me.Label29.Visible = False
Me.lblTanggal.Visible = False
Me.txtTotalBelanja.Visible = False
Me.lblPendapatan.Visible = False
Me.lblLaporan.Visible = False
Me.lblPenjualan.Visible = False
Me.lblJam.Visible = False
Me.Label43.Visible = False
Set ws = Sheets("NOTA")
ws.Activate
Me.tbTransaksi.ColumnCount = 7
Me.tbTransaksi.ColumnHeads = True
Me.tbTransaksi.ColumnWidths = ("30pt; 80 pt;100 pt;0 pt;60 pt;30 pt;60 pt")
Me.tbTransaksi.RowSource = "a7:g8"
Application.Visible = False
Call createNoNota
cetakPenjualan
End Sub
Sub datane()
On Error Resume Next
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Me.lblTerjual.Value = WorksheetFunction.Sum(Sheet4.Range("b2:b" & lastRow))
Me.lblPendapatan.Value = WorksheetFunction.Sum(Sheet4.Range("g2:g" & lastRow))
Me.lblPendapatan.Value = Format(Me.lblPendapatan.Value, "Rp #,##0")
lblTgl.Caption = Date
End Sub
Sub grafike()
On Error Resume Next
Dim ws As Worksheet, ws1 As Worksheet
Set ws = Worksheets("LAPORAN")
Set ws1 = Worksheets("GRAFIK")
ws1.Range("a1").Value = "'" & Date - 6
ws1.Range("a2").Value = "'" & Date - 5
ws1.Range("a3").Value = "'" & Date - 4
ws1.Range("a4").Value = "'" & Date - 3
ws1.Range("a5").Value = "'" & Date - 2
ws1.Range("a6").Value = "'" & Date - 1
ws1.Range("a7").Value = "'" & Date - 0
    ws1.Range("b7").Value = Format(Me.lblPendapatan.Text, "GENERAL NUMBER")
    cari = ws1.Range("a1:a6")
    If Err.Number = 0 Then
    For Each hasil In cari
    Set datanya = ws.Range("b:b").Find(what:=hasil, LookIn:=xlValues)
    Set sumbernya = ws1.Range("a:a").Find(what:=hasil, LookIn:=xlValues)
    baris = sumbernya.Row
    ws1.Cells(baris, 2).Value = datanya.Offset(0, 2)
    baris = baris + 1
    Next hasil
    End If
End Sub
Sub waktu()
On Error Resume Next
Berhenti = False
Do Until Berhenti
    lblJam.Caption = Format(Time, "hh:mm:ss")
    lblTanggal.Caption = WorksheetFunction.Text(Date, "[$-0421] DDDD, DD MMMM YYYY")
    lblTanggalJam.Caption = WorksheetFunction.Text(Date, "[$-0421] DDDD, DD MMMM YYYY") & " - Pkl. " & Format(Time, "hh:mm:ss")
    DoEvents
Loop
End Sub
Sub jamDigital()
On Error Resume Next
Berhenti = False
Do Until Berhenti
    JAM.Caption = Format(Time, "hh:mm:ss")
    TGL.Caption = WorksheetFunction.Text(Date, "[$-0421] DDDD, DD MMMM YYYY")
    DoEvents
Loop
End Sub
Sub cek()
On Error Resume Next
If Sheet6.Range("e1").Value <> Sheet6.Range("e2").Value Then
    If Sheet4.Range("a3").Value = "" Then
        Exit Sub
    End If
Else
Exit Sub
End If
End Sub
Private Sub UserForm_Terminate()
With Application.FileDialog(msoFileDialogFilePicker)
End With
Sheet6.Range("e1").Value = Date
Dim ish As Long
ish = ThisWorkbook.Sheets("DATABARANG").Range("B" & Application.Rows.Count).End(xlUp).Row
Sheets("DATABARANG").Range("A1:m" & ish).Copy Sheets("DATABARANG_BACKUP").Range("A1")
ThisWorkbook.Save
ThisWorkbook.Close
End Sub
