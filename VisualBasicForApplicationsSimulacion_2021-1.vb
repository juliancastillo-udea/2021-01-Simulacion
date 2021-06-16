'Pasos para trabajar convenientemente
'1.	En la pestaña archivo, vamos a opciones, 
'cinta de opciones, en fichas principales 
'activamos la pestaña programador.
'2.	Vamos al centro de confianza y en 
'configuración del centro de confianza, 
'clic en configuración de macros y habilitamos 
'todas las macros y aceptar.
'3.	Clic en archivo, guardar como en grupo de 
'tipo de Excel y seleccionamos libro de Excel 
'habilitado para macros y guardamos.
'4.	Verificar que haya quedado bien guardado.
'5.	Habilitamos ver extensiones de archivo.

'Explorador de proyectos y entorno de programador
'1.	Clic en pestaña programador y el primer 
'botón es Visual Basic
'2.	Clic en Visual Basic y tenemos el explorador 
'de proyectos
'3.	Tenemos dos ventanas en la parte izquierda 
'visibles, tenemos el explorador de proyectos y 
'las propiedades que están anidadas al explorador 
'de proyectos, si están no visibles, clic en ver y 
'seleccionamos ventana de propiedades y explorador de proyectos
'4.	Las hojas tienen dos nombres, uno es la etiqueta 
'del nombre, la que se puede visualizar en Excel 
'directamente que permite espacios, el otro es el
' nombre interno que no permite espacios, es como 
'el nombre de una variable, en la ventana de propiedades 
'este nombre se encuentra entre paréntesis.

'Seleccionar celdas independientes Excel
Range("A3,C3,D7,B10,A6").Select

'Definición de variables en VBA

Sub Ejemplo1()
'
'Definir variables

Dim a01 As Boolean
' Boolean solo permite valores de verdadero y falso en ingles
a01 = True
Range("A1").Select 'Tambien funciona con Cells.
ActiveCell.Value = "Variable de Tipo Booleano"
Range("A2").Select
ActiveCell.Value = a01

Dim a02 As Byte
' Byte solo permite valores enteros muy pequeño entre 0 y 255
a02 = 6
Range("B1").Select
ActiveCell.Value = "Variable de Tipo Byte"
Range("B2").Select
ActiveCell.Value = a02

Dim a03 As Currency
' Currency es para valores de tipo monetario, permite gran tamaño
'  permite rangos desde  
 -922,337,203,685,477.5808 to 922,337,203,685,477.5807
a03 = 526548.3
Range("C1").Select
ActiveCell.Value = "Variable de Tipo Currency o Moneda"
Range("C2").Select
ActiveCell.Value = a03

Dim a04 As Date
' Date permite datos tipo fecha
' Estos datos de tipo fecha van desde el año 
'100 hasta el 31 de diciembre de 9999
' Las horas van desde 
0:00:00 a 23:59:59
a04 = Format(Date, "Short Date")
a04 = "12-12/2012"
Range("D1").Select
ActiveCell.Value = "Variable de Tipo Fecha"
Range("D2").Select
ActiveCell.Value = a04

a04 = Format(Date, "Short Date")
a04 = "12-12-2012"
Range("D3").Select
ActiveCell.Value = a04

a04 = Format(Date, "Short Date")
a04 = "12/12/2012"
Range("D4").Select
ActiveCell.Value = a04

'a04 = Format(Date, "Long Date")
a04 = #5/17/1993 9:32:00 AM#
Range("D5").Select
ActiveCell.Value = a04

Dim a05 As Double
' Double es nuestro tipo Real, permite 
'valores muy grandes con decimal
' El rango va desde 
-1.79769313486231E308 to -4.94065645841247E-324 
'para valores negativos
' Y desde 
4.94065645841247E-324 to 1.79769313486232E308 
'para valores positivos
a05 = 55.2546579846517
Range("E1").Select
ActiveCell.Value = "Variable de Tipo Real - Doble"
Range("E2").Select
ActiveCell.Value = a05

Dim a06 As Integer
' Integer es nuestro tipo entero, 
'permite valores entre 
-32768 a 32767.

a06 = 32767
Range("F1").Select
ActiveCell.Value = "Variable de Tipo Entero"
Range("F2").Select
ActiveCell.Value = a06

Dim a07 As Long
' Long es un tipo entero de tamaño mayor que va desde 
-2147483648 a 2147483647
a07 = 2147483647
Range("G1").Select
ActiveCell.Value = "Variable de Tipo Long"
Range("G2").Select
ActiveCell.Value = a07

Dim a08 As Single
' Single es un tipo de variable de coma flotante 
'que permite valores desde
 -3.402823E38 a -1.401298E-45 'para negativos y desde 
1.401298E-45 a 3.402823E38 
'para positivos
a08 = 2147483647.5
Range("H1").Select
ActiveCell.Value = "Variable de Tipo Single"
Range("H2").Select
ActiveCell.Value = a08

Dim a09 As String
' String es una cadena de texto que puede alojar '
'hasta 2 billones de caracteres.
a09 = "Algoritmos y programacion 6-8"
Range("I1").Select
ActiveCell.Value = "Variable de Tipo String"
Range("I2").Select
ActiveCell.Value = a09

Dim a10 As Variant
' Variant es un tipo de variable que permite 
'todo tipo de datos excepto strings de longitud definida
a10 = 6546879465#
Range("J1").Select
ActiveCell.Value = "Variable de Tipo Variant"
Range("J2").Select
ActiveCell.Value = a10

a10 = "Usar tipo de dato Variant rebaja un punto en el trabajo final"
Range("J3").Select
ActiveCell.Value = a10

 
Columns("A:J").EntireColumn.AutoFit
End Sub
'
'Definir variables


Identificador de caracter	Tipo de Dato	Ejemplo
%							Integer			Dim L%
&							Long			Dim M&
@							Decimal			Const W@ = 37.5
!							Single			Dim Q!
#							Double			Dim X#
$							String			Dim V$ = "Secret"

'------------------------------------------------------------------------
' Operadores Aritmeticos
' Suma +
' Resta -
' Multiplicacion *
' Division /
' Potenciacion ^
' Division entera \
' Modulo Mod

' Operadores Relacionales
' Mayor que >
' Menor que <
' Igual =
' Mayor igual que >=
' Menor igual que <=
' Diferente <>

' Operadores Logicos
' Conjuncion And
' Disyuncion Or
' Negacion Not

'------------------------------------------------------------------------
' Usar y Seleccionar Hojas

Worksheets(1).Select
Cells(1, 1).Value = DateTime.Time
Columns("A:A").ColumnWidth = 25
'Worksheets.Add.Name = "Enero"
Worksheets.Add


Worksheets("La A1").Activate
Sheets("La A1").Select
'------------------------------------------------------------------------
'Valores Generales ------------------------------------------------------------------------------------------------------------------
'Seleccionar Una Hoja
Sheets("Hoja2").Select
'Seleccionar un Rango
Range("A1:A10").Select
' Utilizar la formula aleatorio entre en una celda especifica
Selection.FormulaR1C1 = "=RANDBETWEEN(-10,10)"
'Aplicar formato de texto a una seleccion
With Selection.Font
    .Bold = True
    .Italic = True
End With
'Asignar un valor a un rango seleccionado
Range("A1:A10").Select
ActiveCell.Value = "Total Suma"
'Un cuadro de mensaje sencillo
MsgBox "Hola"
'Utilizar input box
MyNum = Application.InputBox("Enter a number")
Range("A1").Select
ActiveCell.Value = MyNum
'La celda A1 es el origen 1,1 de un plano cartesiano invertido positivamente
'Escribir un la celda A1
Cells(1,1).Value = "Arroz"

'----------------------------------------------------------------------------------------------

' Usar la opcion InputBox

Dim Nombre As String
Nombre = InputBox("Ingresar el nombre del participante." & vbCrLf & "Por defecto se nombrará Goku como participante", "Nombre", "Goku")
Range("D24").Value = Nombre
'------------------------------------------------------------------------
'CONDICIONALES

'1.	If  - Then
'Solo valida si la sentencia es verdadera, 
'no existe cláusula para la parte negativa.
'Ejemplo

Dim dblNota As Double, strResultado As String
dblNota = Range("A1").Value

If dblNota >= 3# Then strResultado = "Materia Coronada"
Range("B1").Value = strResultado
'----------------------------


'2.	If – Thel – Else 
'Esta sentencia contempla las dos clausulas 
'tanto la parte afirmativa como la parte negativa
'Ejemplo

Dim dblNota As Double, strResultado As String
dblNota = Range("A1").Value

If dblNota >= 3# Then 
	strResultado = "Materia Coronada"
Else
	strResultado = "Materia No Coronada"
End

Range("B1").Value = strResultado
'----------------------------

'3.	If – Thel – Else If
'Esta sentencia contempla las dos clausulas tanto 
'la parte afirmativa como la parte negativa y además 
'realiza una nueva validación para el valor negativo.
'Ejemplo

Dim dblNota As Double, strResultado As String
dblNota = Range("A1").Value

If dblNota >= 3# Then 
	strResultado = "Materia Coronada"
ElseIf dblNota <=2# Then
		strResultado = "Materia Muy Perdida, Pailas, nos vemos el otro semestre"
	Else
		strResultado = "Materia Perdida, pero habilitable
End If
Range("B1").Value = strResultado

'------------------------------------------------------------------------

'SELECTOR MULTIPLE Case

'Selector Múltiple Case
'La instrucción Select Case es similar a la instrucción If... Then, ya que pone a prueba una expresión, y lleva a cabo diferentes acciones, dependiendo del valor de la expresión.

'CASO 1
Sub SelectorMultipleSimple()
' Defino los titulos
Range("A1").Value = "Numeros Aleatorio"
Range("B1").Value = "Numero de caso"
' Defino un rango de numeros aleatorio en la columna A
' desde 1 hasta 25
    Range("A2:A26").Select
    Selection.FormulaR1C1 = "=RANDBETWEEN(1,5)"
' Copio y pego valores para no tener formulas dinamicas
    Selection.Copy
    Selection.PasteSpecial _
    Paste:=xlPasteValues '
Range("C1").Select
Selection.Clear
' AutoAjusto las columnas
Columns("A:B").EntireColumn.AutoFit
'Defino el control del ciclo y los rangos
Dim CeldaB As Range
Dim i As Integer
i = 2

' Ciclo For Each obtiene cada uno de los valores del rango
For Each Celda In Range("A2:A26")
' El selector multiple utiliza Celda para definir el valor
    Select Case Celda
   Case 1 ' Caso en el que Celda sea igual a 1
      Range("B" & i).Value = "Caso 1"
   Case 2 ' Caso en el que Celda sea igual a 2
      Range("B" & i).Value = "Caso 2"
   Case 3 ' Caso en el que Celda sea igual a 3
      Range("B" & i).Value = "Caso 3"
   Case 4 ' Caso en el que Celda sea igual a 4
      Range("B" & i).Value = "Caso 4"
   Case 5 ' Caso en el que Celda sea igual a 5
      Range("B" & i).Value = "Caso 5"
' La combinacion de palabras reservadas End Select termina el case
   End Select
    ' Incremento i para imprimir los valores
    i = i + 1
' La palabra reservada Next termina el contenido del ciclo
Next Celda

End Sub

'--------------------------------------------
'CASO 2

Sub SelectorMultipleLetras()
' Defino los titulos
Range("D1").Value = "Letras Aleatorias"
Range("E1").Value = "Letra de caso"
' Defino un rango de letras aleatorioas entre ABCDE en la columna A
' desde 1 hasta 25
    Range("D2:d26").Select
    Selection.FormulaR1C1 = "=CHAR(RANDBETWEEN(65,69))"
' Copio y pego valores para no tener formulas dinamicas
    Selection.Copy
    Selection.PasteSpecial _
    Paste:=xlPasteValues '
Range("C1").Select
Selection.Clear
' AutoAjusto las columnas
Columns("D:E").EntireColumn.AutoFit
'Defino el control del ciclo y los rangos
Dim Celda2 As Range
Dim i As Integer
i = 2

' Ciclo For Each obtiene cada uno de los valores del rango
For Each Celda2 In Range("D2:D26")
' El selector multiple utiliza Celda para definir el valor
    Select Case Celda2
   Case "A" ' Caso en el que Celda sea la letra A
      Range("E" & i).Value = "Caso Letra A"
   Case "B" ' Caso en el que Celda sea la letra B
      Range("E" & i).Value = "Caso Letra B"
   Case "C" ' Caso en el que Celda sea la letra C
      Range("E" & i).Value = "Caso Letra C"
   Case "D" ' Caso en el que Celda sea la letra D
      Range("E" & i).Value = "Caso Letra D"
   Case "E" ' Caso en el que Celda sea la letra E
      Range("E" & i).Value = "Caso Letra E"
' La combinacion de palabras reservadas End Select termina el case
   End Select
    ' Incremento i para imprimir los valores
    i = i + 1
' La palabra reservada Next termina el contenido del ciclo
Next Celda2

End Sub

'--------------------------------------------
'CASO 3
Sub SelectorMultipleRangos()
' Defino los titulos
Range("G1").Value = "Numeros aleatorios"
Range("H1").Value = "Valor Caso"
' Defino un rango de valores aleatorios entre 1 y 100 en la columna G
' desde 1 hasta 25
    Range("G2:G26").Select
    Selection.FormulaR1C1 = "=RANDBETWEEN(1,100)"
' Copio y pego valores para no tener formulas dinamicas
    Selection.Copy
    Selection.PasteSpecial _
    Paste:=xlPasteValues '
Range("I1").Select
Selection.Clear
' AutoAjusto las columnas
Columns("G:H").EntireColumn.AutoFit
'Defino el control del ciclo y los rangos
Dim Celda3 As Range
Dim i As Integer
i = 2

' Ciclo For Each obtiene cada uno de los valores del rango
For Each Celda3 In Range("G2:G26")
' El selector multiple utiliza Celda para definir el valor
    Select Case Celda3
   Case 1 To 20 ' Caso en el que Celda sea entre 1 y 20
      Range("H" & i).Value = "Valor Entre 0 y 20"
   Case 21 To 40 ' Caso en el que Celda sea entre 20 y 40
      Range("H" & i).Value = "Valor Entre 20 y 40"
   Case 41 To 60 ' Caso en el que Celda sea entre 40 y 60
      Range("H" & i).Value = "Valor Entre 40 y 60"
   Case 61 To 80 ' Caso en el que Celda sea entre 60 y 80
      Range("H" & i).Value = "Valor Entre 60 y 80"
   Case 80 To 100 ' Caso en el que Celda sea entre 80 y 100
      Range("H" & i).Value = "Valor Entre 80 y 100"
' La combinacion de palabras reservadas End Select termina el case
   End Select
    ' Incremento i para imprimir los valores
    i = i + 1
' La palabra reservada Next termina el contenido del ciclo
Next Celda3

End Sub

'--------------------------------------------
'CASO 4

Sub SelectorMultipleValoresSingulares()
' Defino los titulos
Range("J1").Value = "Numeros aleatorios"
Range("K1").Value = "Valor Caso"
' Defino un rango de letras aleatorios entre 1 y 10 en la columna j
' desde 1 hasta 25
    Range("J2:J26").Select
    Selection.FormulaR1C1 = "=RANDBETWEEN(1,10)"
' Copio y pego valores para no tener formulas dinamicas
    Selection.Copy
    Selection.PasteSpecial _
    Paste:=xlPasteValues '
Range("l1").Select
Selection.Clear
' AutoAjusto las columnas
Columns("J:K").EntireColumn.AutoFit
'Defino el control del ciclo y los rangos
Dim Celda4 As Range
Dim i As Integer
i = 2

' Ciclo For Each obtiene cada uno de los valores del rango
For Each Celda4 In Range("J2:J26")
' El selector multiple utiliza Celda para definir el valor
    Select Case Celda4
   Case 1, 3, 5 ' Caso en el que Celda sean los valores 1 3 5
      Range("K" & i).Value = "Valores 1, 3, 5"
   Case 2, 4  ' Caso en el que Celda sean los valores 2 4
      Range("K" & i).Value = "Valores 2 y 4"
   Case 6  ' Caso en el que Celda sean los valores 6
      Range("K" & i).Value = "Valor 6"
   Case 7, 8  ' Caso en el que Celda sean los valores 7 8
      Range("K" & i).Value = "Valores 7 y 8"
   Case 9, 10  ' Caso en el que Celda sean los valores 9 10
      Range("K" & i).Value = "Valores 9 y 10"
' La combinacion de palabras reservadas End Select termina el case
   End Select
    ' Incremento i para imprimir los valores
    i = i + 1
' La palabra reservada Next termina el contenido del ciclo
Next Celda4
' AutoAjusto las columnas
Columns("J:K").EntireColumn.AutoFit
Range("L1").Select
Selection.Clear
End Sub

'--------------------------------------------
'CASO 5

Sub SelectorMultipleCondicional()
' Defino los titulos
Range("M1").Value = "Numeros aleatorios"
Range("N1").Value = "Valor Caso"
' Defino un rango de valores aleatorioas entre 1 y 100 en la columna m
' desde 1 hasta 25
    Range("M2:M26").Select
    Selection.FormulaR1C1 = "=RANDBETWEEN(1,100)"
' Copio y pego valores para no tener formulas dinamicas
    Selection.Copy
    Selection.PasteSpecial _
    Paste:=xlPasteValues '
Range("O1").Select
Selection.Clear
' AutoAjusto las columnas
Columns("M:N").EntireColumn.AutoFit
'Defino el control del ciclo y los rangos
Dim Celda5 As Range
Dim i As Integer
i = 2

' Ciclo For Each obtiene cada uno de los valores del rango
For Each Celda5 In Range("M2:M26")
' El selector multiple utiliza Celda para definir el valor
    Select Case Celda5
   Case Is > 80 ' Caso en el que Celda es mayor de 80
      Range("N" & i).Value = "Mayor de 80"
   Case Is > 60  ' Caso en el que Celda es mayor de 60 y menor de 80
      Range("N" & i).Value = "Valores mayor de 60 y menor de 80"
   Case Is > 40  ' Caso en el que Celda es mayor de 40 y menor de 60
      Range("N" & i).Value = "Valores mayor de 40 y menor de 60"
   Case Is > 20  ' Caso en el que Celda es mayor de 20 y menor de 40
      Range("N" & i).Value = "Valores mayor de 20 y menor de 40"
   Case Else ' Caso en el que Celda  es mayor de 0 y menor de 20
      Range("N" & i).Value = "Valores mayor de 0 y menor de 20"
' La combinacion de palabras reservadas End Select termina el case
   End Select
    ' Incremento i para imprimir los valores
    i = i + 1
' La palabra reservada Next termina el contenido del ciclo
Next Celda5
' AutoAjusto las columnas
Columns("M:N").EntireColumn.AutoFit
Range("O1").Select
Selection.Clear
End Sub


'------------------------------------------------------------------------
' Preguntar al usuario
Sub PreguntaUsuario()

' Realizaremos una pregunta en la cual si el usuario responde afirmativamente
' Ejecutaremos una accion de lo contrario no se hará nada

' Creo un string para guardar la pregunta
Dim strPregunta As String
strPregunta = "Desea escribir los numeros del 1 al 100 en la columna A"
' Creo un valor boleano para guardar la respuesta del usuario
Dim blnRespuesta As Boolean
' Verifico la respuesta del usuario y la paso a mi variable booleana
If MsgBox(strPregunta, vbYesNo + vbQuestion) = vbYes Then
    blnRespuesta = True
Else
    blnRespuesta = False
End If
' Creo un rango para imprimir los numeros desde a2 hasta a102
Dim rngRangoImpresion As Range
' Utilizo la palabra "SET" para definir el rango y volverlo definitivo
Set rngRangoImpresion = Range("A2:A102")
' Creo una variable para el ciclo for each
Dim Celdas As Range
If blnRespuesta Then
    For Each Celdas In rngRangoImpresion
    ' CeldaG varia por cada iteracion
    Celdas.Value = i
    ' Incremento i para imprimir los valores
    i = i + 1
    ' La palabra reservada Next termina el contenido del ciclo
    Next Celdas
Else
    For Each Celdas In rngRangoImpresion
    ' CeldaG varia por cada iteracion
    Celdas.Value = Null
    ' Incremento i para imprimir los valores
    i = i + 1
    ' La palabra reservada Next termina el contenido del ciclo
    Next Celdas
    MsgBox "Ha decidido no imprimir los numeros, por lo tanto se borrará toda la informacion", vbInformation, "No Ejecutar"
End If

End Sub


'------------------------------------------------------------------------

'CICLOS
'Ciclo For… Next
'El ciclo For... Next utiliza una variable la cual 
'se fija para cada valor dentro de un rango específico. 
'El código dentro del ciclo se ejecuta entonces para cada valor.

Sub CicloFor()

' Concateno las celdas A1 y B1
Range("A1:B1").Select
Selection.Merge
' Centro el titulo
Selection.HorizontalAlignment = xlCenter
' Escribo titulos
Range("A1").Value = "Ciclo For"
Range("A2").Value = "Numeros del 1 al 10"
Range("B2").Value = "Numeros del 1 al 10"
' AutoAjusto las columnas
Columns("A:B").EntireColumn.AutoFit

'Defino el control del ciclo
Dim i As Integer
' Defino el ciclo For con i como control
' hasta el numero de iteraciones deseadas

For i = 1 To 10
    ' Utilizo dos formas de imprimir valores del ciclo
    Cells(i + 2, 2).Value = i
    Range("A" & i + 2).Value = i
' La palabra reservada Next termina el contenido del ciclo
Next i
End Sub

'--------------------------------------------
'Ciclo Do While
'El ciclo Do While ejecuta repetidamente una sección de código, 
'mientras que una condición especificada continúa evaluada en verdadero (True). 

Sub CicloDoWhile()

' Concateno las celdas D1 y E1
Range("D1:E1").Select
Selection.Merge
' Centro el titulo
Selection.HorizontalAlignment = xlCenter
' Escribo titulos
Range("D1").Value = "Ciclo Do While"
Range("D2").Value = "Numeros del 1 al 10"
Range("E2").Value = "Numeros del 1 al 10"
' AutoAjusto las columnas
Columns("D:E").EntireColumn.AutoFit
'Defino el control del ciclo
Dim i As Integer
' Inicio el ciclo en 1
i = 1
Do While i < 11
    ' Utilizo dos formas de imprimir valores del ciclo
    Cells(i + 2, 4).Value = i
    Range("E" & i + 2).Value = i
    ' Incremento el ciclo de tal manera que pueda tener fin
    i = i + 1
' La palabra reservada Loop termina el contenido del ciclo
Loop

End Sub

'--------------------------------------------
'Ciclo For Each
'El ciclo For Each es similar al ciclo For... Next, 
'pero, en lugar de correr a través de un conjunto de 
'valores para una variable, el ciclo For Each se 
'ejecuta a través de cada objeto dentro de un conjunto de objetos.

Sub CicloForEach()
' Concateno las celdas G1 y H1
Range("G1:H1").Select
Selection.Merge
' Centro el titulo
Selection.HorizontalAlignment = xlCenter
' Escribo titulos
Range("G1").Value = "Ciclo For Each"
Range("G2").Value = "Numeros del 1 al 10"
Range("H2").Value = "Numeros del 1 al 10"
' AutoAjusto las columnas
Columns("G:H").EntireColumn.AutoFit
'Defino el control del ciclo y los rangos
Dim CeldaG As Range
Dim CeldaH As Range
Dim i, j As Integer
i = 1
j = 1

' Primer ciclo For Each imprime cada uno 
'de los valores del rango
For Each CeldaG In Range("G3:G12")
    ' CeldaG varia por cada iteracion
    CeldaG.Value = i
    ' Incremento i para imprimir los valores
    i = i + 1
' La palabra reservada Next termina el contenido del ciclo
Next CeldaG
' Segundo ciclo For Each imprime cada uno de los valores del rango
For Each CeldaH In Range("H3:H12")
    ' CeldaH varia por cada iteracion
    CeldaH.Value = j
    ' Incremento i para imprimir los valores
    j = j + 1
' La palabra reservada Loop termina el contenido del ciclo
Next CeldaH
End Sub

'--------------------------------------------
'Ciclo While
'El ciclo While es similar al ciclo do while, 
'pero, en lugar de evaluar al final la condiciones
'lo hace al principio

Sub CicloWhile()
' Concateno las celdas G1 y H1
Range("J1:K1").Select
Selection.Merge
' Centro el titulo
Selection.HorizontalAlignment = xlCenter
' Escribo titulos
Range("J1").Value = "Ciclo While"
Range("J2").Value = "Numeros del 1 al 10"
Range("K2").Value = "Numeros del 1 al 10"
' AutoAjusto las columnas
Columns("J:K").EntireColumn.AutoFit
'Defino el control del ciclo y los rangos
Dim i as integer
i = 1
While i < 11
    ' Utilizo dos formas de imprimir valores del ciclo
    Cells(i + 2, 4).Value = i
    Range("K" & i + 2).Value = i
    ' Incremento el ciclo de tal manera que pueda tener fin
    i = i + 1
' La palabra reservada Wend termina el contenido del ciclo
Wend

End Sub

'--------------------------------------------
'La instrucción Exit For 
'Si desea salir antes de un ciclo 'For', 
'puede utilizar la instrucción Exit For. 
'Esta declaración provoca que el programa 
'salte fuera del ciclo y continuar con la 
'siguiente línea de código fuera del ciclo. 
'Por ejemplo, usted puede estar buscando un 
'valor determinado en una matriz o vector. 
'Usted puede realizar un bucle a través de 
'cada entrada de la matriz, pero cuando encuentre 
'el valor que está buscando, usted ya no desea 
'continuar la búsqueda, por lo que se sale del 
'ciclo temprano.

Sub CicloExitFor()

' Concateno las celdas A14 y B14
Range("A14:B14").Select
Selection.Merge
' Centro el titulo
Selection.HorizontalAlignment = xlCenter
' Escribo titulos
Range("A14").Value = "Ciclo ExitFor"
Range("A15").Value = "Numeros del 1 al 6"
Range("B15").Value = "Numeros del 1 al 6"
' AutoAjusto las columnas
Columns("A:B").EntireColumn.AutoFit
'Defino el control del ciclo
Dim i As Integer
' Defino el ciclo For con i como control
' hasta el numero de iteraciones deseadas

For i = 1 To 10
    ' Utilizo dos formas de imprimir valores del ciclo
    Cells(i + 15, 2).Value = i
    Range("A" & i + 15).Value = i
    If (i = 6) Then
    ' Si se encuentra el valor 6 el ciclo termina
        Exit For
    End If
' La palabra reservada Next termina el contenido del ciclo
Next i
End Sub

'--------------------------------------------

'El ciclo Do Until 
'El ciclo Do Until es muy similar al Do While. 
'El ciclo Do Until ejecuta repetidamente una sección 
'de código hasta que una condición especificada se evalúa 
'como verdadera (True). Este ciclo valida la información 
'desde su primera iteración o al final, se realiza el 
'ejemplo con las dos condiciones.

Sub CicloDoUntilInicio()
    ' Concateno las celdas D1 y E1
    Range("D14:E14").Select
    Selection.Merge
    ' Centro el titulo
    Selection.HorizontalAlignment = xlCenter
    ' Escribo titulos
    Range("D14").Value = "Ciclo Do Until Inicio"
    Range("D15").Value = "Numeros del 1 al 10"
    Range("E15").Value = "Numeros del 1 al 10"
    ' AutoAjusto las columnas
    Columns("D:E").EntireColumn.AutoFit
    'Defino el control del ciclo
    Dim i As Integer
    ' Inicio el ciclo en 1
    i = 1
    Do Until i = 11
        ' Utilizo dos formas de imprimir valores del ciclo
        Cells(i + 15, 4).Value = i
        Range("E" & i + 15).Value = i
        ' Incremento el ciclo de tal manera que pueda tener fin
        i = i + 1
    ' La palabra reservada Loop termina el contenido del ciclo
    Loop
End Sub

'--------------------------------------------
Sub CicloDoUntilFinal()
    ' Concateno las celdas D1 y E1
    Range("G14:H14").Select
    Selection.Merge
    ' Centro el titulo
    Selection.HorizontalAlignment = xlCenter
    ' Escribo titulos
    Range("G14").Value = "Ciclo Do Until Final"
    Range("G15").Value = "Numeros del 1 al 10"
    Range("H15").Value = "Numeros del 1 al 10"
    ' AutoAjusto las columnas
    Columns("G:H").EntireColumn.AutoFit
    'Defino el control del ciclo
    Dim i As Integer
    ' Inicio el ciclo en 1
    i = 1
    Do
        ' Utilizo dos formas de imprimir valores del ciclo
        Cells(i + 15, 7).Value = i
        Range("H" & i + 15).Value = i
        ' Incremento el ciclo de tal manera que pueda tener fin
        i = i + 1
    ' La palabra reservada Loop termina el contenido del ciclo
    Loop Until i = 11
End Sub



'------------------------------------------------------------------------

'PROTEGER Y DESPROTEGER LIBROS Y HOJAS

ActiveWorkbook.Unprotect ("Batman")
Sheets("Notas").Unprotect Password:="Batman"

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="Batman"
ActiveWindow.SelectedSheets.Visible = False

ActiveWorkbook.Protect ("Batman")
'-------------------------------------------
' PROTEGER LIBRO Y HOJA
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="Batman"
ActiveWorkbook.Protect ("Batman")

'N---------------------------------------------------------------------------------------------------------------------------------------------------------------

'Encontrar el ultimo registro de una columna N---------------------------------------------------------------------------------------
	Dim Ultimo As Double
	Ultimo = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Cells(Ultimo + 1, 1).Value = "Dato1"
    Cells(Ultimo + 1, 2).Value = "Dato2"

'Encontrar el ultimo registro de una columna N---------------------------------------------------------------------------------------



'Uso usuario y contraseña -------------------------------------------------------------------------------------------------------
Private Sub btnIngreso_Click()
    ' Defino y obtengo la informacion del usuario
    Dim NombreUsuario As String
    NombreUsuario = txtUsuario.Text
    ' Defino y obtengo la informacion de la contraseña
    Dim Contrasenia As String
    Contrasenia = txtContrasenia.Text
    ' Verifico que si se ha ingresado la informacion de usuario
    If IsNull(Me.txtUsuario) Or Me.txtUsuario = "" Then
        MsgBox "Favor ingresar informacion de usuario.", vbOKOnly, "INFORMACION REQUERIDA"
        Me.txtUsuario.SetFocus
        Exit Sub
    End If
     
    ' Verifico si se ha ingresado datos en la informacion de contraseña
    If IsNull(Me.txtContrasenia) Or Me.txtContrasenia = "" Then
        MsgBox "Favor ingresar la constraseña.", vbOKOnly, "INFORMACION REQUERIDA"
        Me.txtContrasenia.SetFocus
        Exit Sub
    End If
     
    ' Verficamos que el usuario y la contraseña estan en la base de datos registrados para acceder

    '.WorksheetFunction.VLOOKUP(valor_a_buscar, rango, indice_columna, Falso)
    On Error GoTo UsuarioNoExiste:
    If (NombreUsuario = WorksheetFunction.VLookup(NombreUsuario, Seguridad.Range("A1:B20"), 1, 0)) Then

        If (Contrasenia = WorksheetFunction.VLookup(NombreUsuario, Seguridad.Range("A1:B20"), 2, 0)) Then
            MsgBox "Contrasenia & Nombre de Usuario Validos", vbOKOnly, "INGRESO ACEPTADO"
            Unload Me 'Cierro el formulario actual
            frmInicio.Show ' Muestro el formulario destino
        Else
            MsgBox "Nombre de Usuario & Contraseña no validos, vuelva a intentarlo", vbOKOnly, "INGRESO NO ACEPTADO"
        End If
    Else
UsuarioNoExiste:
        MsgBox "Nombre de usuario no existe", vbOKOnly, "INGRESO NO ACEPTADO"
    End If
End Sub
'Uso usuario y contraseña -------------------------------------------------------------------------------------------------------

'Abrir por defecto un formulario al inicio de excel------------------------------------------------------------------------------
' Esto se hace en el workbook
Private Sub Workbook_Open()
formulario.Show
End Sub
'Abrir por defecto un formulario al inicio de excel------------------------------------------------------------------------------

'Llenar Combos ------------------------------------------------------------------------------------------------------------------
'Opcion 1 Llenar ComboBox
   Me.cboBarrio.List = Worksheets("Combos").Range("B2:B372").Value

' Opcion 2 Llenar ComboBox
   With Worksheets("Combos")
       cboBarrio.List = .Range("B2:B" & .Range("B" & .Rows.Count).End(xlUp).Row).Value
   End With

'Opcion 3 Llenar ComboBox
    Dim rng As Range
    For Each rng In Worksheets("Combos").Range("B2:B372")
       cboBarrio.AddItem rng.Value
    Next
    cboBarrio.ListIndex = 0
'Llenar Combos ------------------------------------------------------------------------------------------------------------------

' Vectores y matrices------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------
'Vectores y Matrices
Public Sub UsoOffset()
    'Utilizamos el comando With para crear un entorno de propiedades
    'Seleccionamos la Hoja1 del libro actual para trabajar con ella
    With ThisWorkbook.Worksheets("Hoja1")
        ' Se declaran variables para estudiantes
        Dim Titulo As String
        Dim Estudiante1 As Integer
        Dim Estudiante2 As Integer
        Dim Estudiante3 As Integer
        Dim Estudiante4 As Integer
        Dim Estudiante5 As Integer

        ' Leemos las notas desde excel
        Titulo = .Range("A1").Offset(0)
        Estudiante1 = .Range("A1").Offset(1)
        Estudiante2 = .Range("A1").Offset(2)
        Estudiante3 = .Range("A1").Offset(3)
        Estudiante4 = .Range("A1").Offset(4)
        Estudiante5 = .Range("A1").Offset(5)
        ' Ver - Ventana Inmediato
        ' Debug.Print imprime en la ventana de salida de comando
        Debug.Print "Notas estudiantes"
        Debug.Print Estudiante1
        Debug.Print Estudiante2
        Debug.Print Estudiante3
        Debug.Print Estudiante4
        Debug.Print Estudiante5
    End With
End Sub

'------------------------------------------------------------------------
Public Sub Definir()
    With ThisWorkbook.Worksheets("Hoja1")
        ' Definicion de Variable
        Dim Estudiante As Integer
        Dim Pais As String
        ' Definicion de Vector
        Dim Estudiantes(1 To 3) As Integer
        Dim Paises(1 To 3) As String
        ' Asignacion de datos a las variables
        ' Asignacion a la variable
        Estudiante = .Cells(1, 1)

        ' Asignacion de valor al vector en la posicion 1
        Estudiantes(1) = .Cells(1, 1)
        ' Impresion de la variable
        Debug.Print Estudiante
        ' Impresion del vector en la posicion 1
        Debug.Print Estudiantes(1)
    End With
End Sub
'------------------------------------------------------------------------
Public Sub DefinicionVectores()
    'Utilizamos el comando With para crear un entorno de propiedades
    'Seleccionamos la Hoja1 del libro actual para trabajar con ella
    ' Para visualizar la salida imediata seleccionamos ver y luego Ventana Inmediato
    With ThisWorkbook.Worksheets("Hoja1")
        ' Declaracion de vectores para notas de estudiantes
        Dim Estudiantes(1 To 5) As Integer
        ' Leemos las notas del ejercicio anterior desde la casilla A2 hacia adelante
        Dim i As Integer
        For i = 1 To 5
            Estudiantes(i) = .Range("A1").Offset(i)
        Next i
        ' Impresion en consola inmediato de la inforamcion
        Debug.Print "Students Marks"
        For i = LBound(Estudiantes) To UBound(Estudiantes)
            Debug.Print Estudiantes(i)
        Next i
    End With
End Sub
'------------------------------------------------------------------------
Public Sub VectoresEstaticos()

    ' Crear un vector con datos desde 0 a 3
    Dim VectorTipo1(0 To 3) As Long
	' Crear un vector con datos desde 0 a 3
    Dim VectorTipo2(3) As Long
    ' Crear un vector con posiciones de 1 a 5 / 0 no esta presente
    Dim VectorTipo3(1 To 5) As Long
    ' Crear un vector con posciones de 2 a 4 / muy poco usado
    Dim VectorTipo4(2 To 4) As Long
End Sub
'------------------------------------------------------------------------
Public Sub VectoresDinamicos()

    ' Declaracion de un vector dinamico, su tamaño no se encuentra definido
	Dim VectorDinamico() As Long
	' Se define la extension del Vector cuando se necesita
    ReDim VectorDinamico(0 To 5)
	' La cantidad maxima de elementos en un vector en excel es aproximadamente 65mil
	' Al pasar este limite sale el siguiente error
	'Out of Memory
	'1005: Unable to set the Value property of the Range class 
End Sub
'------------------------------------------------------------------------
Public Sub AsignarValores()
    ' Definir un vector con valores de 0 a 3
    Dim AsignarValoresVector(0 To 3) As Long
    ' Asignacion de un valor en la posicion 0
    AsignarValoresVector(0) = 5
    ' Asignacion de un valor en la posicion 3
    AsignarValoresVector(3) = 46
    ' Esto es un error ya que no hay posicion 4
    AsignarValoresVector(4) = 99

End Sub
'------------------------------------------------------------------------
Public Sub Dividir()
Dim Vector1 As Variant
Vector1 = Array("Orange", "Peach", "Pear")
Dim Vector2 As Variant
Vector2 = Array(5, 6, 7, 8, 12)
'La funcion Array crea un vector de indice cero
' La funcion split divide texto o datos de acuerdo a un parametro de separacion
' Un parametro de separacion puede ser coma u otro tipo caracter.
' Crear un string para guardar la inforamcio a dividir
Dim Dividir As String
' Separado por comas para poder realizar la separacion
Dividir = "Red,Yellow,Green,Blue"
' Creamos un Vector Dinamico para realizar la separacion
Dim VectorDinamico() As String
' Usamos la funcion split para realizar la separacion
VectorDinamico = Split(Dividir, ",")
End Sub
'------------------------------------------------------------------------
Public Sub ArrayLoops()
    ' Definimos un Vector
    Dim VectorM(0 To 5) As Long
    ' Llenamos el vector con numeros aleatorios
    Dim i As Long
    ' Usando las funciones Lbound y UBound, las cuales
    ' Toman los valores limite inferior y superior de los vectores
    For i = LBound(VectorM) To UBound(VectorM)
        VectorM(i) = 10 * Rnd
    Next i
    ' Impresion de los valores mediante debug
    Debug.Print "Ubicacion", "Valor"
    For i = LBound(VectorM) To UBound(VectorM)
        Debug.Print i, VectorM(i)
    Next i
End Sub
'------------------------------------------------------------------------
Public Sub UsandoForEach()
	'Definimos un vector cualquiera
    Dim VectorM(0 To 5) As Long

    ' Llenamos el vector con numeros aleatorios|
    Dim i As Long
    For i = LBound(VectorM) To UBound(VectorM)
        VectorM(i) = 100 * Rnd
    Next i
    ' Definimos un dato tipo variant
    Dim Dato As Variant
	'Utilizamos el for each para recorrer el vector e imprimir el dato
    For Each Dato In VectorM
        Debug.Print Dato
    Next Dato
End Sub
'------------------------------------------------------------------------
Public Sub BorrarVector()
    ' Creamos un vector cualquiera
    Dim Vector(0 To 3) As Long
    ' Llenamos el vector de datos aleatorios
    Dim i As Long
    For i = LBound(Vector) To UBound(Vector)
        Vector(i) = 5 * Rnd
    Next i
    ' Limpiar el vector facil y suave
    Erase Vector
    ' impresion de los valores que ahora son dtodos cero
    Debug.Print "Localizacion", "Valor"
    For i = LBound(Vector) To UBound(Vector)
        Debug.Print i, Vector(i)
    Next i
End Sub
'------------------------------------------------------------------------

Public Sub Matrices()

    ' Definimos un vector bidimensional
    Dim Matriz(0 To 3, 0 To 2) As String
    'El vector es de 4 filas y 3 columnas
    
    ' Llenamos el vectro con datos creados por funciones
    Dim i As Long, j As Long
    For i = LBound(Matriz) To UBound(Matriz)
        For j = LBound(Matriz, 2) To UBound(Matriz, 2)
        'La funcion CStr convierte un dato a tipo String
            Matriz(i, j) = CStr(i) & ":" & CStr(j)
        Next j
    Next i

    ' Realizamos la impresion de los valores
    Debug.Print "i", "j", "Valores"
    For i = LBound(Matriz) To UBound(Matriz)
        For j = LBound(Matriz, 2) To UBound(Matriz, 2)
            Debug.Print i, j, Matriz(i, j)
        Next j
    Next i
End Sub
'------------------------------------------------------------------------
Public Sub UsandoForEachMatriz()
' Usamos el mismo codigo anterior para crear la matriz y agregar los valores
    ' Definimos un vector bidimensional
    Dim Matriz(0 To 3, 0 To 2) As String
    'El vector es de 4 filas y 3 columnas
    
    ' Llenamos el vectro con datos creados por funciones
    Dim i As Long, j As Long
    For i = LBound(Matriz) To UBound(Matriz)
        For j = LBound(Matriz, 2) To UBound(Matriz, 2)
        'La funcion CStr convierte un dato a tipo String
            Matriz(i, j) = CStr(i) & ":" & CStr(j)
        Next j
    Next i
    
    ' usando  For se necesitan dos ciclos
    Debug.Print "i", "j", "Value"
    For i = LBound(Matriz) To UBound(Matriz)
        For j = LBound(Matriz, 2) To UBound(Matriz, 2)
            Debug.Print i, j, Matriz(i, j)
        Next j
    Next i

    ' Usando ForEach es solo un ciclo
    Debug.Print "Valor"
    Dim Dato As Variant
    For Each Dato In Matriz
        Debug.Print Dato
    Next Dato
End Sub

'------------------------------------------------------------------------
Public Sub ReadToArray()

    ' Declare dynamic array
    Dim StudentMarks As Variant

    ' Read values into array from first row
    StudentMarks = Range("A1:Z1").Value

    ' Write the values back to the third row
    Range("A3:Z3").Value = StudentMarks

End Sub
'------------------------------------------------------------------------
Public Sub ReadAndDisplay()

    ' Get Range
    Dim rg As Range
    Set rg = ThisWorkbook.Worksheets("Sheet1").Range("C3:E6")

    ' Create dynamic array
    Dim StudentMarks As Variant

    ' Read values into array from sheet1
    StudentMarks = rg.Value

    ' Print the array values
    Debug.Print "i", "j", "Value"
    Dim i As Long, j As Long
    For i = LBound(StudentMarks) To UBound(StudentMarks)
        For j = LBound(StudentMarks, 2) To UBound(StudentMarks, 2)
            Debug.Print i, j, StudentMarks(i, j)
        Next j
    Next i

End Sub

' Vectores y matrices------------------------------------------------------------------------------------------------------------

' Funciones ---------------------------------------------------------------------------------------------------------------------
' Para definir una funcion se utiliza la palabra reservada "Function" luego el nombre, diferente a los nombres de las funciones de excel
' Luego abre parentesis y define los parametros de entrada, tantos como necesite, hasta 255
' Luego es opcional definir el tipo de la funcion como se define una variable
' Es super importante que se establezca un resultado, el objetivo de la funcion es retornar un resultado
' El resultado es una variable del mismo nombre con un valor del mismo tipo que la expresion de funcion.
' Ejemplo generico de funcion
Function NombreDeFuncionAqui (Parametro1 as string, Parametro2 as integer, Parametro3 as double) as string

' Ejemplo de validacion de 3 tipos de datos para una funcion 
If Parametro1 = "A" then
	If Parametro2 = 3 then
		If Parametro3 = 5 then
			NombreDeFuncionAqui = "El parametro 1 es A, el parametro 2 es 3 y el parametro 3 es 5"
		Else
			NombreDeFuncionAqui = "El parametro 1 es A, el parametro 2 es 3 y el parametro 3 no es 5"
		End If
	Else
		NombreDeFuncionAqui = "El parametro 1 es A, el parametro 2 no es 3 y el parametro 3 no es 5"
	End If
Else
	NombreDeFuncionAqui = "El parametro 1 no es A, el parametro 2 es 3 y el parametro 3 no es 5"
End If

End Function


Function Suma_xy(x As Double, y As Double) As Double

Suma_xy = x + y

End Function

Function Multiplicacion_xy(x As Double, y As Double) As Double

Multiplicacion_xy = x * y

End Function

Function Division_xy(x As Double, y As Double) As Double

Division_xy = x / y

End Function

Function Resta_xy(x As Double, y As Double) As Double

Resta_xy = x - y

End Function

' Ahora algunos un poco mas complejos
' Realizar un ejercicio donde la suma se divida con la resta
' y se multiplique con la multiplicacion

Function Hardcore() As Double

Hardcore = Division_xy((Suma_xy(Cells(2, 1).Value, Cells(2, 2).Value)), (Resta_xy(Cells(2, 1).Value, Cells(2, 2).Value))) * Multiplicacion_xy(Cells(2, 1).Value, Cells(2, 2).Value)

End Function

Function Potenciacion(x As Double, y As Double) As Double
Potenciacion = x ^ y
End Function

Function Radicacion(x As Double, y As Double) As Double
Radicacion = x ^ (1 / y)
End Function

Function AleatorioEntreAB(A As Integer, B As Integer) As Integer
'A = -100
'B = 100
'Formula Aleatorio Entre dos numeros mediante VBA sin usar RANDBETWEEN
' Bueno que guarden esta formula
AleatorioEntreAB = Int(Rnd() * (B - A + 1) + A)
End Function

' Funciones ---------------------------------------------------------------------------------------------------------------------
