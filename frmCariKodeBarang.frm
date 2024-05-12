VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CARIKODEBARANG 
   Caption         =   "CARI KODE BARANG"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "frmCariKodeBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CARIKODEBARANG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tbBarang_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
TAMBAHSTOK.cboKodeBarang.Value = Me.tbBarang.Column(1)
TAMBAHSTOK.txtStokMasuk.SetFocus
Unload Me
End Sub
Private Sub txtCari_Change()
On Error Resume Next
    If Me.txtCari.Value = "" Then
    With Sheets("DATABARANG")
    .AutoFilterMode = False
    End With
    tampil
    Else
    Sheets("DATABARANG").Activate
    Application.ScreenUpdating = False
    Sheets("DATABARANG").Range("A1:G1").AutoFilter Field:=3, Criteria1:="*" & Me.txtCari.Text & "*", Operator:=xlAnd
    Sheets("HASILFILTER").Cells.clear
    Sheets("DATABARANG").Range("A1:g" & Cells(Rows.Count, 1).End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy _
    Destination:=Sheets("HASILFILTER").Range("A1")
    Sheets("HASILFILTER").Cells.EntireColumn.AutoFit
    Sheets("HASILFILTER").Select
    Me.tbBarang.RowSource = "HASILFILTER!a2:g" & Cells(Rows.Count, 1).End(xlUp).Row + 1
    Application.ScreenUpdating = True
    End If
End Sub
Sub tampil()
On Error Resume Next
Set ws = Sheets("DATABARANG")
ws.Activate
akhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    Me.tbBarang.RowSource = "a2:g" & akhir + 1
End Sub
Private Sub UserForm_Initialize()
On Error Resume Next
Set ws = Sheets("DATABARANG")
ws.Activate
akhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    Me.tbBarang.RowSource = "a2:g" & akhir + 1
Me.tbBarang.ColumnWidths = ("0pt;50pt;180pt;0pt;0pt;0pt;60pt")
txtCari.SetFocus
End Sub
