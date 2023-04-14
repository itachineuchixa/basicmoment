VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "������������ ��������"
   ClientHeight    =   2010
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5280
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnu_file 
      Caption         =   "����"
      Begin VB.Menu mnu_open 
         Caption         =   "�������"
      End
      Begin VB.Menu mnu_close 
         Caption         =   "�������"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ��������
Dim xlapp As Excel.Application
Dim xlbook As Excel.Workbook
Dim xlsheet As Excel.Worksheet
'�������������
'Dim Prepxlapp As Excel.Application
Dim Prepxlbook As Excel.Workbook
Dim Prepxlsheet As Excel.Worksheet


Private Sub mnu_open_Click()
Dim pyt As String
Dim kollst As Integer
Dim i As Integer
Dim j As Integer
Dim skip As Integer
Dim prepod As String
Dim grup As String
CommonDialog1.DialogTitle = "������� ��������"
CommonDialog1.CancelError = False
CommonDialog1.Filter = "*.xls"
CommonDialog1.ShowOpen
pyt = CommonDialog1.FileName
If pyt = "" Then Exit Sub
Set xlapp = CreateObject("Excel.Application")
Set xlbook = xlapp.Workbooks.Open(pyt)
xlapp.Workbooks.Open (App.Path & "\������ �����.xls")
xlapp.Workbooks.Open (App.Path & "\������ ����-�������.xls")
xlapp.Workbooks.Open (App.Path & "\������ �������.xls")
Set xlsheet = xlapp.Workbooks(1).Worksheets(1)
xlsheet.Activate
kollst = xlbook.Worksheets.Count
ProgressBar1.Max = kollst + 3
ProgressBar1.Value = 1
ProgressBar1.Visible = True
DoEvents
'xlapp.Visible = True
For i = 1 To kollst
	Try
		ProgressBar1.Value = ProgressBar1.Value + 1
		Set xlsheet = xlbook.Worksheets(i)
		Select Case xlsheet.Range("A6")
		
		'**********************
			Case Is = "����-������� ���������"
			j = 11 ' ������ ������������ ���������
				While xlsheet.Cells(j, 2) <> Chr(34) & "�����������" & Chr(34)
					' ������ ������
					If xlsheet.Range("X" & j) <> "" Then
						prepod = Trim(xlsheet.Range("X" & j))
						grup = xlsheet.Name
						Set xlsheet = xlapp.Workbooks(3).Worksheets(1)
						xlsheet.Activate
						'���������� ������ �� ���� �� �������
						If provPrepodList(prepod, 3) = 0 Then
						' �� ������- ������
							xlapp.Workbooks(3).Sheets("����-�������").Select
							xlapp.Workbooks(3).Sheets("����-�������").Copy After:=xlapp.Workbooks(3).Sheets(1)
							xlapp.Workbooks(3).Sheets("����-������� (2)").Select
							xlapp.Workbooks(3).Sheets("����-������� (2)").Name = prepod
							'����������� � ������� ������
							Set xlbook = xlapp.Workbooks(1)
							Set xlsheet = xlbook.Sheets(i)
							xlsheet.Activate
							 Range("B" & j & ":W" & j).Select
							Selection.Copy
							Windows("������ ����-�������.xls").Activate
							Set xlsheet = xlapp.Workbooks(3).Sheets(2)
							xlsheet.Activate
							Range("B11").Select
							xlsheet.Paste
							'�������� ��� ������� � excel
							xlsheet.Range("A3").Value = xlsheet.Range("A3").Value + " " & prepod
							' ����� ������
							xlsheet.Range("C11").Value = grup
						Else
							' ������ - ��������� � �����
							  '����������� � ������� ������
							Set xlbook = xlapp.Workbooks(1)
							Set xlsheet = xlbook.Sheets(i)
							xlsheet.Activate
							 Range("B" & j & ":W" & j).Select
							Selection.Copy
							Windows("������ ����-�������.xls").Activate
							Set xlsheet = xlapp.Workbooks(3).Sheets(prepod)
							xlsheet.Activate
							Range("B" & poisk_Stroki(prepod, 3)).Select
							xlsheet.Paste
							 ' ����� ������ -1 ������ ��� 2-�� ���
							xlsheet.Range("C" & poisk_Stroki(prepod, 3) - 1).Value = grup

						End If
						Set xlsheet = xlapp.Workbooks(1).Sheets(i)
						xlsheet.Activate
					Else
						' ������ �� ������
						List1.AddItem xlsheet.Name & "  " & xlsheet.Range("B" & j) & "  �� ������ �������������"
					End If
					j = j + 1
				Wend
	'*************************
			Case Is = "����� ���������"
	'*************************
				j = 11 ' ������ ������������ ���������
				While xlsheet.Cells(j, 2) <> Chr(34) & "�����������" & Chr(34)
					' ������ ������
					If xlsheet.Range("X" & j) <> "" Then
						prepod = Trim(xlsheet.Range("X" & j))
						grup = xlsheet.Name
						Set xlsheet = xlapp.Workbooks(2).Worksheets(1)
						xlsheet.Activate
						'���������� ������ �� ���� �� �������
						If provPrepodList(prepod, 2) = 0 Then
						' �� ������- ������
							xlapp.Workbooks(2).Sheets("�����").Select
							xlapp.Workbooks(2).Sheets("�����").Copy After:=xlapp.Workbooks(2).Sheets(1)
							xlapp.Workbooks(2).Sheets("����� (2)").Select
							xlapp.Workbooks(2).Sheets("����� (2)").Name = prepod
							'����������� � ������� ������
							Set xlbook = xlapp.Workbooks(1)
							Set xlsheet = xlbook.Sheets(i)
							xlsheet.Activate
							 Range("B" & j & ":W" & j).Select
							Selection.Copy
							Windows("������ �����.xls").Activate
							Set xlsheet = xlapp.Workbooks(2).Sheets(2)
							xlsheet.Activate
							Range("B11").Select
							xlsheet.Paste
							'�������� ��� ������� � excel
							xlsheet.Range("A3").Value = xlsheet.Range("A3").Value + " " & prepod
							' ����� ������
							xlsheet.Range("C11").Value = grup
						Else
							' ������ - ��������� � �����
							  '����������� � ������� ������
							Set xlbook = xlapp.Workbooks(1)
							Set xlsheet = xlbook.Sheets(i)
							xlsheet.Activate
							 Range("B" & j & ":W" & j).Select
							Selection.Copy
							Windows("������ �����.xls").Activate
							Set xlsheet = xlapp.Workbooks(2).Sheets(prepod)
							xlsheet.Activate
							Range("B" & poisk_Stroki(prepod, 2)).Select
							xlsheet.Paste
							 ' ����� ������ -1 ������ ��� 2-�� ���
							xlsheet.Range("C" & poisk_Stroki(prepod, 2) - 1).Value = grup

						End If
						Set xlsheet = xlapp.Workbooks(1).Sheets(i)
						xlsheet.Activate
					Else
						' ������ �� ������
						List1.AddItem xlsheet.Name & "  " & xlsheet.Range("B" & j) & "  �� ������ �������������"
					End If
					j = j + 1
				Wend
	'*************************
			Case Is = "������� ���������"
			 j = 11 ' ������ ������������ ���������
				While xlsheet.Cells(j, 2) <> Chr(34) & "�����������" & Chr(34)
					' ������ ������
					If xlsheet.Range("M" & j) <> "" Then
						prepod = Trim(xlsheet.Range("M" & j))
						grup = xlsheet.Name
						Set xlsheet = xlapp.Workbooks(4).Worksheets(1)
						xlsheet.Activate
						'���������� ������ �� ���� �� �������
						If provPrepodList(prepod, 4) = 0 Then
						' �� ������- ������
							xlapp.Workbooks(4).Sheets("�������").Select
							xlapp.Workbooks(4).Sheets("�������").Copy After:=xlapp.Workbooks(4).Sheets(1)
							xlapp.Workbooks(4).Sheets("������� (2)").Select
							xlapp.Workbooks(4).Sheets("������� (2)").Name = prepod
							'����������� � ������� ������
							Set xlbook = xlapp.Workbooks(1)
							Set xlsheet = xlbook.Sheets(i)
							xlsheet.Activate
							 Range("B" & j & ":L" & j).Select
							Selection.Copy
							Windows("������ �����.xls").Activate
							Set xlsheet = xlapp.Workbooks(4).Sheets(2)
							xlsheet.Activate
							Range("B11").Select
							xlsheet.Paste
							'�������� ��� ������� � excel
							xlsheet.Range("A3").Value = xlsheet.Range("A3").Value + " " & prepod
							' ����� ������
							xlsheet.Range("C11").Value = grup
						Else
							' ������ - ��������� � �����
							  '����������� � ������� ������
							Set xlbook = xlapp.Workbooks(1)
							Set xlsheet = xlbook.Sheets(i)
							xlsheet.Activate
							 Range("B" & j & ":L" & j).Select
							Selection.Copy
							Windows("������ �������.xls").Activate
							Set xlsheet = xlapp.Workbooks(4).Sheets(prepod)
							xlsheet.Activate
							Range("B" & poisk_Stroki(prepod, 4)).Select
							xlsheet.Paste
							 ' ����� ������ -1 ������ ��� 2-�� ���
							xlsheet.Range("C" & poisk_Stroki(prepod, 4) - 1).Value = grup

						End If
						Set xlsheet = xlapp.Workbooks(1).Sheets(i)
						xlsheet.Activate
					Else
						' ������ �� ������
						List1.AddItem xlsheet.Name & "  " & xlsheet.Range("B" & j) & "  �� ������ �������������"
					End If
					j = j + 1
				Wend
			Case Else
				List1.AddItem xlsheet.Name & " � ������ �6 �� ������� ����� �������� "
		End Select
	Next i
	DoEvents
	itogo 2
	itogo 3
	itogo 4
	xlapp.Workbooks(2).SaveAs App.Path & "\�����.xls"
	xlapp.Workbooks(3).SaveAs App.Path & "\����-�������.xls"
	xlapp.Workbooks(4).SaveAs App.Path & "\�������.xls"
	xlapp.Workbooks(4).Close
	xlapp.Workbooks(3).Close
	xlapp.Workbooks(2).Close
	xlapp.Workbooks(1).Close

	Set xlapp = Nothing
	Set xlbook = Nothing
	Set xlsheet = Nothing
	ProgressBar1.Visible = False
	MsgBox "�������� ������������", vbInformation, "�������"
	End Sub
	Private Function provPrepodList(prepod As String, book As Integer) As Integer
	Dim kollst As Integer
	Dim i As Integer
	'Windows("������ �����.xls").Activate
	kollst = xlapp.Workbooks(book).Worksheets.Count
	Set xlsheet = xlapp.Workbooks(book).Worksheets(1)
	provPrepodList = 0
	For i = 1 To kollst
		Set xlsheet = xlapp.Workbooks(book).Worksheets(i)
		If xlsheet.Name = prepod Then provPrepodList = i
	Next i
	End Function
	Private Function poisk_Stroki(prepod As String, book As Integer) As Integer
	Dim kollst As Integer
	Dim i As Integer
	Set xlsheet = xlapp.Workbooks(book).Worksheets(provPrepodList(prepod, book))
	xlsheet.Activate
	i = 11
	poisk_Stroki = 0
	While xlsheet.Range("B" & i).Value <> ""
		i = i + 1
	Wend
	poisk_Stroki = i

	End Function
	'ActiveCell.FormulaR1C1 = "=SUM(R[-6]C:R[-2]C)"
	Private Sub itogo(book As Integer)
	Dim kollst As Integer
	Dim nom As Integer
	Dim i As Integer
	kollst = xlapp.Workbooks(book).Sheets.Count
	ProgressBar1.Max = kollst + 3
	ProgressBar1.Value = 1
	For i = 2 To kollst
		ProgressBar1.Value = ProgressBar1.Value + 1
		Set xlsheet = xlapp.Workbooks(book).Sheets(i)
		xlsheet.Activate
		nom = poisk_Stroki(xlsheet.Name, book)
		'xlsheet.Range("B" & nom + 1).Select
		'ActiveCell.FormulaR1C1 = "=SUM(R[-" & nom - 11 & "]C:R[-2]C)"
		xlsheet.Range("B" & nom + 1).Value = "�����:"
		xlsheet.Range("B" & nom + 1).Select
		With Selection
			.HorizontalAlignment = xlRight
			.VerticalAlignment = xlBottom
			.WrapText = False
			.Orientation = 0
			.AddIndent = False
			.IndentLevel = 0
			.ShrinkToFit = False
			.ReadingOrder = xlContext
			.MergeCells = False
			.Font.Bold = True
		End With
		' �������
		If book = 4 Then
			' �� �����
			xlsheet.Range("K" & nom + 1).FormulaR1C1 = "=SUM(R[-" & nom - 10 & "]C:R[-2]C)"
			xlsheet.Range("K" & nom + 1).Font.Bold = True
			' �� �������������
			xlsheet.Range("L" & nom + 1).FormulaR1C1 = "=SUM(R[-" & nom - 10 & "]C:R[-2]C)"
			xlsheet.Range("L" & nom + 1).Font.Bold = True
			' �����
			xlsheet.Range("B" & nom + 2) = "�����"
			xlsheet.Range("B" & nom + 2).Font.Bold = True
			xlsheet.Range("K" & nom + 2) = CInt(xlsheet.Range("D" & nom + 1)) + CInt(xlsheet.Range("M" & nom + 1))
			xlsheet.Range("K" & nom + 2).Font.Bold = True
		'���������
		Else
			' �� �����
			xlsheet.Range("D" & nom + 1).FormulaR1C1 = "=SUM(R[-" & nom - 10 & "]C:R[-2]C)"
			xlsheet.Range("D" & nom + 1).Font.Bold = True
			' �� �������������
			xlsheet.Range("W" & nom + 1).FormulaR1C1 = "=SUM(R[-" & nom - 10 & "]C:R[-2]C)"
			xlsheet.Range("W" & nom + 1).Font.Bold = True
			' �����
			xlsheet.Range("B" & nom + 2) = "�����"
			xlsheet.Range("B" & nom + 2).Font.Bold = True
			xlsheet.Range("D" & nom + 2) = CInt(xlsheet.Range("D" & nom + 1)) + CInt(xlsheet.Range("W" & nom + 1))
			xlsheet.Range("D" & nom + 2).Font.Bold = True
		End If
		
	Catch ex As Exception
		skip = 1
Next i

End Sub
