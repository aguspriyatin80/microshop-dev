VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CARICUSTOMER 
   Caption         =   "CARI CUSTOMER"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6765
   OleObjectBlob   =   "frmCariCustomer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CARICUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tbCustomer_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
DASHBOARD2.txtCustomer.Value = Me.tbCustomer.Column(1) & " - " & Me.tbCustomer.Column(3)
DASHBOARD2.txtCustomer.SetFocus
Unload Me
End Sub
Private Sub txtCari_Change()
On Error Resume Next
    If Me.txtCari.Value = "" Then
    With Sheets("DATACUSTOMER")
    .AutoFilterMode = False
    End With
    tampil
    Else
    Sheets("DATACUSTOMER").Activate
    Application.ScreenUpdating = False
    Sheets("DATACUSTOMER").Range("A1:D1").AutoFilter Field:=3, Criteria1:="*" & Me.txtCari.Text & "*", Operator:=xlAnd
    Sheets("FILTERCUSTOMER").Cells.clear
    Sheets("DATACUSTOMER").Range("A1:D" & Cells(Rows.Count, 1).End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy _
    Destination:=Sheets("FILTERCUSTOMER").Range("A1")
    Sheets("FILTERCUSTOMER").Cells.EntireColumn.AutoFit
    Sheets("FILTERCUSTOMER").Select
    Me.tbCustomer.RowSource = "FILTERCUSTOMER!a2:g" & Cells(Rows.Count, 1).End(xlUp).Row + 1
    Application.ScreenUpdating = True
    End If
End Sub
Sub tampil()
On Error Resume Next
Set ws = Sheets("DATA_CUSTOMER")
ws.Activate
akhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    Me.tbCustomer.ColumnCount = 4
    Me.tbCustomer.ColumnWidths = ("60pt;100pt;180pt;100pt")
    Me.tbCustomer.RowSource = "a2:d" & akhir + 1
End Sub
Private Sub UserForm_Initialize()
Call tampil
End Sub
