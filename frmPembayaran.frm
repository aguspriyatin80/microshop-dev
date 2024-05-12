VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PEMBAYARAN 
   Caption         =   "PEMBAYARAN"
   ClientHeight    =   3345
   ClientLeft      =   8040
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmPembayaran.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "PEMBAYARAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub animasi()
Dim i As Integer
Dim j As Integer
Dim X As Integer
Dim Y As Integer
PEMBAYARAN.Left = 400
PEMBAYARAN.Top = 0
PEMBAYARAN.BackColor = vbRed
X = 200
Application.Wait Now + TimeValue("00:00:01")
For i = 1 To X
    If PEMBAYARAN.Top < X Then
    PEMBAYARAN.Top = PEMBAYARAN.Top + 1
    Else
    PEMBAYARAN.Top = PEMBAYARAN.Top
    End If
Next i
End Sub
Sub animasi2()
Dim i As Integer
For i = 1 To 400
    If PEMBAYARAN.Top < 400 Then
        'Application.Wait Now + TimeValue("00:00:01")
        PEMBAYARAN.Top = PEMBAYARAN.Top + i
    Else
        Unload Me
    End If
Next i
End Sub
Private Sub cmdFinish_Click()
    animasi2
    Sheets("NOTA").Activate
    DASHBOARD2.selesai
End Sub
Private Sub txt_bayar_AfterUpdate()
Me.txt_bayar.Text = Replace(Me.txt_bayar.Text, ",", "")
End Sub
Private Sub txt_bayar_Change()
Me.txt_bayar.Text = Format(Me.txt_bayar.Text, "#,##0")
Me.txt_kembali.Value = Val(Replace(Me.txt_bayar.Text, ",", "")) - Val(Replace(Me.txt_total_belanja.Text, ",", ""))
Me.txt_kembali.Text = Format(Me.txt_kembali.Text, "#,##0")
DASHBOARD2.txtBayar.Text = Me.txt_bayar.Text
DASHBOARD2.txtKembali.Text = Me.txt_kembali.Text
End Sub
Private Sub UserForm_Activate()
Me.Left = 400
Me.txt_total_belanja.Value = DASHBOARD2.lblTotalBelanja.Caption
Me.txt_diskon.Value = 0
animasi
End Sub
Private Sub UserForm_Initialize()
Me.txt_bayar.SetFocus
End Sub
