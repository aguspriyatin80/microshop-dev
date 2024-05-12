VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PROFILTOKO 
   Caption         =   "PROFIL TOKO"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   OleObjectBlob   =   "frmProfilToko.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PROFILTOKO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim uploadImage As Boolean
Dim lokasi_foto As String
Sub simpan()
Dim lokasi_foto As String
On Error Resume Next
DASHBOARD2.lblNamaToko.Caption = Me.txtNamaToko.Text
DASHBOARD2.lblAlamat1.Caption = Me.txtAlamat1.Text
DASHBOARD2.lblAlamat2.Caption = Me.txtAlamat2.Text
Sheets("PROFIL_TOKO").Range("C1").Value = Me.txtNamaToko.Text
Sheets("PROFIL_TOKO").Range("C2").Value = Me.txtAlamat1.Text
Sheets("PROFIL_TOKO").Range("C3").Value = Me.txtAlamat2.Text
Sheets("NOTA").Range("C1").Value = UCase(Me.txtNamaToko.Text)
    ActiveSheetUsedRange = Sheets("NOTA").Range("c2:g2")
    For Each c In ActiveSheetUsedRange
        If c.WrapText Then
            c.Rows.AutoFit
            Rows("2:2").RowHeight = 25
        End If
    Next c
Sheets("NOTA").Range("C2").Value = UCase(Me.txtAlamat1.Text)
Sheets("NOTA").Range("C3").Value = UCase(Me.txtAlamat2.Text)
If uploadImage Then
    If Application.FileDialog(msoFileDialogFilePicker).SelectedItems.Count > 0 Then
        With Application.FileDialog(msoFileDialogFilePicker)
            lokasi_foto = .SelectedItems(1)
        End With
        Sheets("PROFIL_TOKO").Range("C4").Value = lokasi_foto
        DASHBOARD2.logoToko.Picture = LoadPicture(Sheets("PROFIL_TOKO").Range("C4").Value)
        Me.imgLogoToko.Picture = LoadPicture(Sheets("PROFIL_TOKO").Range("C4").Value)
        Set ws = Sheets("NOTA")
        DeleteImage
        ws.Range("d1").Select
        ws.Pictures.Insert(lokasi_foto).Select
        selection.ShapeRange.Width = 20
        selection.ShapeRange.Height = 20
        Rows("1:1").RowHeight = 35
    End If
Else
        DASHBOARD2.logoToko.Picture = LoadPicture(Sheets("PROFIL_TOKO").Range("C4").Value)
        Me.imgLogoToko.Picture = LoadPicture(Sheets("PROFIL_TOKO").Range("C4").Value)
End If
ThisWorkbook.Save
MsgBox "Profile updated successfully!"
End Sub
Sub DeleteImage()
Dim pic As Picture
ActiveSheet.Unprotect
For Each pic In ActiveSheet.Pictures
    If Not Application.Intersect(pic.TopLeftCell, Range("A1:H2")) Is Nothing Then
        pic.Delete
    End If
Next pic
End Sub
Private Sub cmdRemoveLogo_Click()
uploadImage = False
imgLogoToko.Picture = LoadPicture(vbNullString)
lokasi_foto = ""
DeleteImage
Sheets("PROFIL_TOKO").Range("C4").Value = lokasi_foto
End Sub
Private Sub cmdSimpan_Click()
MsgBox "Maaf, tidak bisa update profil toko!"
End Sub
Private Sub cmdUpload_Click()
uploadImage = True
With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .Filters.Add "Foto", "*.jpg;*.jpeg"
If .Show = -1 Then
imgLogoToko.Picture = LoadPicture(.SelectedItems(1))
lokasi_foto = .SelectedItems(1)
End If
End With
End Sub
Private Sub imgLogoToko_Click()
MsgBox Application.FileDialog(msoFileDialogFilePicker).SelectedItems.Count
End Sub
Private Sub UserForm_Initialize()
Me.txtNamaToko.Text = DASHBOARD2.lblNamaToko.Caption
Me.txtAlamat1.Text = DASHBOARD2.lblAlamat1.Caption
Me.txtAlamat2.Text = DASHBOARD2.lblAlamat2.Caption
Me.imgLogoToko.Picture = LoadPicture(Sheets("PROFIL_TOKO").Range("C4").Value)
End Sub
