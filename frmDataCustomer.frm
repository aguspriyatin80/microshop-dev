VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DATACUSTOMER 
   Caption         =   "DATA CUSTOMER"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmDataCustomer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DATACUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersihkan()
Me.txtNamaCustomer.Text = ""
Me.txtAlamat.Text = ""
Me.txtNoHP.Text = ""
kodeOtomatis
Me.txtNamaCustomer.SetFocus
Me.cmdSimpan.Caption = "SIMPAN"
End Sub
Private Sub cboIDCustomer_Change()
Dim rg As Range
Dim cIndex As Variant
Dim kode As String
On Error Resume Next
lastRow = Sheets("DATA_CUSTOMER").Cells(Rows.Count, 1).End(xlUp).Row
Set rg = Sheets("DATA_CUSTOMER").Range("A2:B" & lastRow)
kode = Me.txtID.Value
cIndex = Application.Match(kode, rg.Columns(1), 1)
If Len(cIndex) > 0 Then
    Me.txtNamaCustomer.Text = rg.Cells(cIndex, 2).Value
End If
End Sub
Sub kodeOtomatis()
Set aktifSheet = Sheets("DATA_CUSTOMER")
lastRow = aktifSheet.Cells(Rows.Count, 2).End(xlUp).Row + 1
kode_otomatis = "CUS" & Format(Right(aktifSheet.Cells(lastRow - 1, 1), 3) + 1, "0##")
Me.txtID.Text = kode_otomatis
End Sub
Private Sub cmdBatal_Click()
bersihkan
Me.cmdHapus.Enabled = False
End Sub
Private Sub cmdHapus_Click()
Dim baris As Long
Dim dbarea As Range, hc As Range
Set aktifSheet = Sheets("DATA_CUSTOMER")
baris = aktifSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
Set dbarea = aktifSheet.Range("A:A" & baris)
Set hc = dbarea.Find(Me.txtID.Text, , xlValues, xlWhole)
If Not hc Is Nothing Then
    hc.EntireRow.Delete
End If
End Sub
Private Sub cmdSimpan_Click()
Dim baris As Integer
If Me.cmdSimpan.Caption = "SIMPAN" Then
    Set aktifSheet = Sheets("DATA_CUSTOMER")
    lastRow = aktifSheet.Cells(Rows.Count, 2).End(xlUp).Row + 1
    If Me.txtNamaCustomer.Text <> "" Then
        aktifSheet.Cells(lastRow, 1) = Me.txtID.Text
        aktifSheet.Cells(lastRow, 2) = txtNamaCustomer
        aktifSheet.Cells(lastRow, 3) = txtAlamat
        aktifSheet.Cells(lastRow, 4) = txtNoHP
    Else
        MsgBox "Nama customer tidak boleh kosong!"
        Me.txtNamaCustomer.SetFocus
    End If
Else
    Set aktifSheet2 = Sheets("DATA_CUSTOMER").Range("A:A")
    Set cari = aktifSheet2.Find(Me.txtID.Value, LookAt:=xlWhole)
        If Not cari Is Nothing Then
            baris = cari.Row
            'aktifSheet2.Cells(baris, 1) = Me.txtID
            aktifSheet2.Cells(baris, 2) = Me.txtNamaCustomer
            aktifSheet2.Cells(baris, 3) = Me.txtAlamat
            aktifSheet2.Cells(baris, 4) = Me.txtNoHP
        End If
End If
tampilkanTabel
bersihkan
Me.cmdHapus.Enabled = False
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.ListBox1.ListIndex = -1 Then Exit Sub
Me.txtID.Value = Me.ListBox1.Column(0)
Me.txtNamaCustomer.Value = Me.ListBox1.Column(1)
Me.txtAlamat.Value = Me.ListBox1.Column(2)
Me.txtNoHP.Value = Me.ListBox1.Column(3)
Me.cmdHapus.Enabled = True
Me.cmdSimpan.Caption = "UPDATE"
End Sub
Private Sub UserForm_Activate()
kodeOtomatis
Me.txtID.Enabled = False
tampilkanTabel
End Sub
Private Sub UserForm_Initialize()
Set aktifSheet = Sheets("DATA_CUSTOMER")
aktifSheet.Select
lastRow = aktifSheet.Cells(Rows.Count, 2).End(xlUp).Row + 1
tampilkanTabel
Me.cmdHapus.Enabled = False
End Sub
Sub tampilkanTabel()
Set aktifSheet = Sheets("DATA_CUSTOMER")
aktifSheet.Select
Me.ListBox1.ColumnCount = 4
Me.ListBox1.ColumnHeads = True
Me.ListBox1.ColumnWidths = "70;90;120;80"
lastRow = Sheets("DATA_CUSTOMER").Cells(Rows.Count, 1).End(xlUp).Row
Me.ListBox1.RowSource = "A2:d" & lastRow
End Sub
