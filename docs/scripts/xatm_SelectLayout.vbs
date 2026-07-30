-----------------------
Documentação de Scripts
-----------------------
xatm_SelectLayout
Thu Jul 30 12:04:47 2026
-----------------------

<xatm_SelectLayout.btnPrevious:btnPrevious_Click()>
Sub btnPrevious_Click()

	Dim imageList
	imageList = xatm_SelectLayout.ImageList

	Dim index
	index = xatm_SelectLayout.Index

	If Trim(imageList) = "" Then Exit Sub

	Dim parts
	parts = Split(imageList, ",")

	Dim layouts()
	ReDim layouts(UBound(parts))

	Dim count
	count = 0

	Dim current
	current = -1

	Dim i
	For i = 0 To UBound(parts)
		If Trim(parts(i)) <> "" Then
			layouts(count) = Trim(parts(i))
			If StrComp(layouts(count), Trim(index), vbTextCompare) = 0 Then current = count
			count = count + 1
		End If
	Next

	If count = 0 Then Exit Sub

	Dim previous
	previous = current - 1
	If previous < 0 Then previous = count - 1

	xatm_SelectLayout.Index = layouts(previous)

End Sub


<xatm_SelectLayout.btnNext:btnNext_Click()>
Sub btnNext_Click()

	Dim imageList
	imageList = xatm_SelectLayout.ImageList

	Dim index
	index = xatm_SelectLayout.Index

	If Trim(imageList) = "" Then Exit Sub

	Dim parts
	parts = Split(imageList, ",")

	Dim layouts()
	ReDim layouts(UBound(parts))

	Dim count
	count = 0

	Dim current
	current = -1

	Dim i
	For i = 0 To UBound(parts)
		If Trim(parts(i)) <> "" Then
			layouts(count) = Trim(parts(i))
			If StrComp(layouts(count), Trim(index), vbTextCompare) = 0 Then current = count
			count = count + 1
		End If
	Next

	If count = 0 Then Exit Sub

	Dim nextOne
	nextOne = current + 1
	If nextOne > count - 1 Then nextOne = 0

	xatm_SelectLayout.Index = layouts(nextOne)

End Sub

