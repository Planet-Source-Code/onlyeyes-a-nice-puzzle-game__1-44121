Attribute VB_Name = "Kernel"
Global theMatrix(8, 8) As Byte
Global pozX, pozY As Byte
Global playerName As String
Global pathDir As String
Type myType
    x As Byte
    y As Byte
End Type

'***********************************************************************
' Procedura initMatrix realizeaza initializarea matricii theMatrix cu
'  valori consecutive in functie de paramentrul primit mode care poare
'  lua valoarea 1 pentru modul beginer si 2 pentru modul advanced.
'***********************************************************************
Sub initMatrix(mode As Byte)
    Dim i, j As Byte
    Dim k As Integer
    
    If mode = 1 Then
        k = -1
        For i = 1 To 4
            For j = 1 To 4
                k = k + 1
                theMatrix(i, j) = k
            Next j
        Next i
    ElseIf mode = 2 Then
        k = -1
        For i = 1 To 8
            For j = 1 To 8
                k = k + 1
                theMatrix(i, j) = k
            Next j
        Next i
    End If
End Sub

'***********************************************************************
' Procedura refreshMatrix realizeaza rescrierea tuturor elementelor cu
'  valoarea 0.Procedura s-a creat din motive de siguranta rolul ei in
'  mare parte fiind indeplinit de procedura initMatrix.
'***********************************************************************
Sub refreshMatrix()
    Dim i, j As Byte
    For i = 1 To 8
        For j = 1 To 8
            theMatrix(i, j) = 0
        Next j
    Next i
End Sub

'***********************************************************************
' Proceduta MatrixPos realizeaza determinarea a doua pozitii x si y in
'  matrice in functie de un numar primit ca parametru si de modul in
'  in care se afla jucatorul. Valorile determinate se vor stoca in
'  variabilele pozX ,pozY care sunt variabile globale cu vizibilitate
'  totala.
'***********************************************************************
Sub MatrixPos(ByVal Number As Byte, mode As Byte)
    Dim i, j, k As Byte
        i = 0
        j = 1
        k = 0
    For i = 1 To Number
        k = k + 1
        If i Mod 4 * mode = 0 Then
            If i < Number Then
                j = j + 1
                k = 0
            End If
        End If
    Next i
    pozX = j
    pozY = k
End Sub

'***********************************************************************
' Functia CreateIndex realizeaza transformarea a doua valori x ,y care
'  reprezinta linia si coloana curenta a matricii theMatrix intr-o
'  singura valoare reprezentand numarul de index al acestui element in
'  matrice in functie de parametrul mode.Va returna aceasta valoare.
'************************************************************************
Function CreateIndex(ByVal x As Byte, ByVal y As Byte, mode As Byte) As Byte
    Dim i, j, k As Byte
    k = 0
    For i = 1 To 4 * mode
        For j = 1 To 4 * mode
            k = k + 1
            If i = x And j = y Then
                CreateIndex = k - 1
            End If
        Next j
    Next i
End Function

'***********************************************************************
' Procedura shuffleMatrix realizeaza amestecul elementelor din matricea
'  the Matrix ,dupa criteriul jocului de Puzzle , in functie de
'  parametrul primit mode.
'***********************************************************************
Sub shuffleMatrix(mode As Byte)
    Dim i, j As Byte
    Dim k As Integer
    Dim poz As Byte
    Dim val As Byte
    Dim tempMatrixVal As Byte
    Dim possible(1 To 4) As myType
    Dim trace(1) As myType
    If mode = 1 Then
        refreshMatrix
        initMatrix (1)
        poz = 0
        i = 4
        j = 4
        trace(1).x = 0
        trace(1).y = 0
        For k = 1 To 100
            If (i - 1 > 0 And i - 1 <> trace(1).x) Then
                    poz = poz + 1
                    possible(poz).x = i - 1
                    possible(poz).y = j
            End If
            If (i + 1 < 5 And i + 1 <> trace(1).x) Then
                    poz = poz + 1
                    possible(poz).x = i + 1
                    possible(poz).y = j
            End If
            If (j - 1 > 0 And j - 1 <> trace(1).y) Then
                    poz = poz + 1
                    possible(poz).x = i
                    possible(poz).y = j - 1
            End If
            If (j + 1 < 5 And j + 1 <> trace(1).y) Then
                    poz = poz + 1
                    possible(poz).x = i
                    possible(poz).y = j + 1
            End If
            Randomize
            val = Int(poz * Rnd + 1)
            tempMatrixVal = theMatrix(possible(val).x, possible(val).y)
            theMatrix(possible(val).x, possible(val).y) = theMatrix(i, j)
            theMatrix(i, j) = tempMatrixVal
            trace(1).x = i
            trace(1).y = j
            i = possible(val).x
            j = possible(val).y
            poz = 0
        Next k
    ElseIf mode = 2 Then
        refreshMatrix
        initMatrix (2)
        poz = 0
        i = 8
        j = 8
        trace(1).x = 0
        trace(1).y = 0
        For k = 1 To 200
            If (i - 1 > 0 And i - 1 <> trace(1).x) Then
                    poz = poz + 1
                    possible(poz).x = i - 1
                    possible(poz).y = j
            End If
            If (i + 1 < 9 And i + 1 <> trace(1).x) Then
                    poz = poz + 1
                    possible(poz).x = i + 1
                    possible(poz).y = j
            End If
            If (j - 1 > 0 And j - 1 <> trace(1).y) Then
                    poz = poz + 1
                    possible(poz).x = i
                    possible(poz).y = j - 1
            End If
            If (j + 1 < 9 And j + 1 <> trace(1).y) Then
                    poz = poz + 1
                    possible(poz).x = i
                    possible(poz).y = j + 1
            End If
            Randomize
            val = Int(poz * Rnd + 1)
            tempMatrixVal = theMatrix(possible(val).x, possible(val).y)
            theMatrix(possible(val).x, possible(val).y) = theMatrix(i, j)
            theMatrix(i, j) = tempMatrixVal
            trace(1).x = i
            trace(1).y = j
            i = possible(val).x
            j = possible(val).y
            poz = 0
        Next k
    End If
       
End Sub

'***********************************************************************
' Procedura whenWin determina cand jucatorul a reusit sa asambleze toate
'  elementele Puzzel-lui si semnaleaza acest lucru printr-un mesaj
'  vizual.
'***********************************************************************
Sub whenWin(mode As Byte)
    Dim i, j As Byte
        For i = 1 To 4 * mode
            For j = 1 To 4 * mode
                If theMatrix(i, j) <> CreateIndex(i, j, mode) Then
                    Exit Sub
                End If
            Next j
        Next i
   MsgBox "You won!.Tray again by presing F5.Tanck's for playng Puzzle Game", vbInformation + vbOKOnly
End Sub

'***********************************************************************
' Functia obtainPath returneaza o valoare de tip string reprezentand
'  calea spre o anumita imagine iar in parametrul fileName care este
'  o referinta stocheaza numele imaginii respective. Parametrul path
'  reprezinta calea + numelefisierului.
'***********************************************************************
Function obtainPath(path As String, ByRef fileName As String) As String
    Dim i As Integer
        For i = Len(path) To 1 Step -1
            If Mid(path, i, 1) = "\" Then
                obtainPath = Mid(path, 1, i - 1)
                fileName = Mid(path, i + 1, Len(path))
                Exit For
            End If
        Next i
End Function
