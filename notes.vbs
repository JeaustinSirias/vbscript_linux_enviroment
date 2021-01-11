' APUNTES Y BASES EN EL USO DE VBScript
' Jeustin Sirias (C) 2021
'
'
'
' COMENTARIOS
' Los comentarios se pueden digitar colocando una comilla 
' simple "'" o bien usando REM
'
'
'
' TIPOS DE DATOS Y VARIABLES
' Solo existe un tipo general de datos 'Variant' y puede 
' contener letras, cifras o booleanos. Por ejemplo,
'
variable  = "This is a char string"
age = 50
estado = true
'
' Para usar una variable primero debe DECLARARSE y luego,
' INICIALIZARSE.
'
DIM user      ' Declarando una variable (NULL)
user = "Juan" ' Inicializando una variable
'
' Para garantizar la calidad de la estructura del codigo 
' es una buena practica obligar al programa a que verifique
' si toda variable ha sido inicializada antes de ser decla-
' rada. Usar OPTION EXPLICIT en el encabezado garantiza esto.
'
'OPTION EXPLICIT
DIM valor 
valor = 5000
'
' 
'
' MATRICES Y VECTORES
' Por la relacion entre Microsoft Excel y VB, es comun nominar
' las matrices como 'tablas'. Igual que en Python y C/C++ se
' puede acceder a los datos de una matriz n x m, por indexacion.
' Para dimensionar la matriz se declara una variable primero
'
DIM matrix (3) ' Indica que un vector 1 x 3 ha sido reservado
'
' Los espacios reservados en 'matrix' pueden ser reservados de
' la siguiente manera:
'
matrix (0) = 10
matrix (1) = 20
matrix (2) = "Pepe"
matrix (3) = false
'
' El arreglo matricial soporta letras, cifras y booleanos en sus
' elementos de su composicion. Cada elemento puede visualizarse 
' usando MSGBOX para imprimir el resultado en pantalla
'
MSGBOX(matrix (2))
'
' Las matrices multidimensionales tambien son admitidas hasta 60
' dimensiones y se declaran de la siguiente forma:
'
DIM arreglo (2, 2, 3) ' Declarar una matriz 3D 
'
' 
'
' ENTRADA Y SALIDA DE DATOS POR PANTALLA Y TECLADO
' VBScript es un lenguaje orientado a la POO y un documento activo
' en pantalla es un objeto en VB llamado DOCUMENT y tiene el metodo
' WRITE para ingresar datos como usuario. Para obtener texto en pan-
' talla se usa la instruccion MSGBOX tambien.
'
' DOCUMENT.WRITE("Ejemplo de impresion") ' No sirve muy bien!
'
' Para ingresar datos por teclado al programa se usa INPUTBOX y tiene
' por argumentos INPUTBOX(texto, titulo, x, y).
'
DIM input
input =  INPUTBOX ("Digite su edad", "Registro", 0, 0)
MSGBOX ("La edad ingresada es " + input) ' Se usa '+' para concatenar
'
'
'
' SOBRE CONDICIONALES IF/THEN/ELSE
' A diferencia de Python y C/C++, los condicionales en VBS requieren 
' una 'confimacion' con la palabra reservada THEN:
'
DIM edad 
edad = INPUTBOX ("Digite su edad", "Registro de ingreso", 0, 0)
IF (edad < 18) THEN 
    MSGBOX ("Usted es menor de edad")
ELSE 
    MSGBOX ("Siga adelante. Bienvenido")
END IF ' Los condicionales IF siempre se deben cerrar con un END IF
'
' Los condicionales anidados tambien son validos en VBS y multiples
' capas pueden ser usadas pruedentemente. Los operadores logicom
' AND y OR tambien son validos en la sintaxis de VBS. El switch/case
' tambien es posible en VBS:
'
DIM estado
estado = INPUTBOX ("1 para ordenar pollo; 2 para ordenar papas", "Menu", 0, 0)
SELECT CASE estado
    CASE 1:
        MSGBOX ("Orden de pollo en progreso")
    CASE 2:
        MSGBOX ("Orden de papas en progreso")
END SELECT
'
'
'
' BUCLES FOR/NEXT Y WHILE/LOOP
' Las iteraciones para los contadores de un  bucle for depende de una
' variable que debe ser declarada previamente (i.e., DIM var) y con-
' trario a lo tradicional en Python y C/C++, el bucle se debe cerrar
' usando la sentencia NEXT.
'
DIM counter
FOR counter = 0 to 5
    'DOCUMENT.WRITE ("HOLA!")
NEXT
'
' Para determinar el orden de 'brincos' en el bucle, se puede utilizar
' la intruccion STEP del siguiente modo.
'
FOR counter = 0 to 20 STEP 2
    'DOCUMENT.WRITE ("Hola!")
NEXT
'
' De forma similar los bucles WHILE/LOOP funcionan. Por ejemplo suponer
' la autenticacion de una clave por ingreso por teclado:
'
DIM clave
DO WHILE (clave <> "usuario123") ' Alternativamente puede usarse UNTIL 
    clave = INPUTBOX ("Digite la clave secreta", "Autenticacion", 0, 0)
LOOP : MSGBOX ("Clave correcta")
'
'
'
' FUNCIONES DE EJECUCION Y PROCEDIMIENTOS
' - Funciones inmediatas
' - Funciones de usuario
' - Funciones del interprete
' Las funciones de ejecucion inmediata son instrucciones programadas
' en forma de procedimientos que no han sido previamente elaborados. 
' De igual manera que en otros lenguajes una funcion puede ejecutarse
' multiples veces en el flujo de ejecucion. 
'
' Funcion suma
FUNCTION suma(A, B)
    suma =  CLNG(A) + CLNG(B) ' CLNG() evita solo concatenar A y B
END FUNCTION

' Declarar variables
DIM A
DIM B
DIM resultado

A = 0 : B = 0

DO UNTIL (A > 0 AND B > 0)
    A = INPUTBOX("Digite el primer numero, A", "Sumador", 0, 0)
    B = INPUTBOX("Digite el segundo numero, B", "Sumador", 0, 0)
LOOP
MSGBOX ("El resultado es: " & suma(A, B))
'
' En lenguajes como C/C++ existen las funciones void, conocidas
' como funciones 'sin retorno'. En VBS estas funciones se deno-
' minan 'Procedimientos' y se encierran entre las instrucciones
' SUB y END SUB. La idea de los procedimientos es solo ejecutar 
' acciones. Para llamar un procedimiento se usa CALL.
'
SUB estados(bool)
    IF bool = true THEN
        MSGBOX ("El valor es verdadero")
    ELSE 
        MSGBOX ("El valor es falso")
    END IF
END SUB
'
CALL estados(false)
'
' 
'
' FUNCIONES DEL LENGUAJE (INTERPRETE)
' Son funciones predefinidas en el inteprete de VBS 
 

