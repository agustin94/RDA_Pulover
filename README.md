#  Que es Pulover Macro Creator?
Es una herramienta de automatización y generador de scripts. Se basa en el lenguaje AutoHotkey y proporciona a los usuarios múltiples funciones de automatización, así como un grabador incorporado.
Sitio: https://www.macrocreator.com/

## Objetivo
Completar la documentacion de word con los datos extraidos del excel, una vez completada la documentacion, enviarlo por mail al cliente correspondiente que figura en el excel.

# Paso a paso

# Excel

## Open Excel.
Con el botón Run/Ejecutar, seleccionamos el excel desde **target** y en el command seleccionamos **run**

![](./gif/RUN_EXCEL.gif)


## Connect macro with Excel.
Para que nuestra macro pueda extraer los datos del excel, tenemos que relacionarlo siguiendo estos pasos.

![](./gif/Conexion_excel_Cominterface.gif)

Steps:
  1. Vamos a Funciones(Functions).
  2. En Variable Output escribiremos **XL**.
  3. En Function Name seleccionamos **ComObjActive**.
  4. En Parámetros escribimos (con comillas) **"Excel. Aplicación** y presionamos aceptar.
  5. Vamos al boton **COM INTERFACE**.
  6. Click en check 'Automatically Create COM OBJECT'
  7. En 'Handle' escribimos **XL**.
  8. En 'CLSID' buscamos la opcion **"Excel.Application"**
  9. Ppresione el botón de conexión. Pulover está minimizado, tienes que buscar el excel que quieres usar, al encontrarlo presiona click 
  derecho (Como podemos ver en el gif), al presionarlo debe salir de una ventana que dice **Connection Succesful!**. 
  Si falla intenta conectarlo de nuevo o verifica que el excel no este dañado.
  10. En 'Expression/COM INTERFACE' Ponemos el valor de la celda y lo guardamos en una variable.
     
  En este ejemplo, toma el valor de una sola celda. Lo guarda dentro de una variable llamada **Dato**.
     
     
     ` Dato := XL.Range("B2").Value `
     
       
  -**XL**: Es el nombre con el que se relaciona con el excel.
  -**Range("Columna y NumeroDeFila")**:    
  -**Value**: Obtenga el valor que tiene dentro de la celda.
  -**Dato**: Es el nombre de la variable, puede poner el nombre que desee y contendrá el valor de la celda.
  Para llamar el valor de una variable lo usamos como un porcentaje% Dato%
     
  Para tomar el valor de varias celdas tenemos que usar ** Copiar **, Pulover guarda los valores copiados en ** Portapapeles **.
  Luego se invoca otra acción con el signo de porcentaje **% Portapapeles% ** (así como una variable):.
     
     ` XL.Range("A2:B4").Copy` 
         
   
   11. Presione el botón 'Ok'.
   
   ##### Opcional: si desea ver el valor, puede ir a **Message Box** y escribir el nombre de la variable de este modo %Nombre de variable%
   
   

## Pausa
Usamos pausa en caso de que la aplicacion tenga alguna demora en alguna accion, nos servira para que el codigo no saltee 
ningun proceso que necesitemos y falle. Nos servira para que el proceso que automatizamos vaya a una velocidad sincronizada
con la aplicacion que estamos manipulando.

Gif pausa
 

  
## WinActivate
Para que la ventana de nuestro aplicativo este siempre activa usamos **WinActivate**.

![](./gif/Win_Activate_Excel.gif)

# WORD

 ## Abrir Word
 Abrimos el word del mismo modo que el excel (es valido para cualquier aplicacion)
 ![](./gif/RUN_WORD.gif)
 
 ## Llamar funcion de word con Atajo de tecla.
 Microsoft Word tiene una opcion para reemplazar palabras, que podemos llamarlo con una **atajo del teclado** escribiendo *CTRL+L*.
 COLOCAR IMAGEN DE MUESTRA
 
 Pulover tiene una opcion para que puedas usar una combinacion de teclado como vimos anteriormente. 
 Pulsamos las teclas que queremos usar en nuestra macro.
 GIF TECLA
 
 Los atajos de tecla nos servira para agilizar nuestra automatizacion y funcione con menos demora.
 Si queres ver mas de estos atajos visita este sitio: https://www.computerhope.com/shortcut/word.htm

 
 
 _______________________________________________________________
 **Click a boton con screenshot**
 REEMPLAZARLO
 ![](./gif/Screenshot_boton.gif)
 
 ## Insertar acciones de teclado.
 Insertamos un atajo de teclado, en este caso usamos Eliminar y la cantidad de veces que desea que se presione.
 ![](./gif/pulse_keyboard.gif)
 

 ## Escribe Texto
 Podemos llenar un campo con una cadena texto que escribiremos usando el boton **TEXT** de pulover. Escribiremos un texto
 o podemos usar una variable que declaramos antes, llamandola para que escriba el texto que contiene esa  dicha variable.
 **Escribiendo un texto**
 
 
 ![](./gif/write_text.gif)
  _______________________________________________________________
  
 **Llamar a una variable que contiene los datos de una celda de Excel solicitada previamente**
 
 **Recordatorio**: *No te olvides que todas las variables que usaremos en un Texto o incluso dentro de otra varible siempre 
 tienen que estar adentro del simbolo de porcentaje %Nombre de la variable%*
 
 
 ![](./gif/Write_Text_Variable.gif)
 
 
 ## Editar y guardar word con datos de excel.
 Con los paso que vimos anteriormente podemos reemplazar las palabras de un texto en Word con valores desde un excel.
 
 Pasos:
 1. Abrimos excel
 2. Conectamos macro a excel.
 3. Escribimos los valores que deseamos extraer en una variable. 
 Ej: `Nombre := XL.Range("B2").Value `
 4. Abrimos Word.
 5. Una vez abierto el word, usamos el atajo de tecla *CTRL+L* y abre la ventana **Buscar y reemplazar**
 6. Escribiremos un texto usando el boton **Text** que vimos anteriormente y escribiremos el boton que queremos reemplazar en word.
 7. En el atajo de teclado escribiremos tab, asi podremos pasar al siguiente campo donde usariamos el texto que queremos usar.
 8. Usaremos enter.
 9. Cuando el word esta completo, lo guardamos sin reemplazar el archivo original. 
 10. Con F12 ingresamos a la ventana *Guardar Como...*
 11. Podemos escribir el nombre del archivo y ruta, como veremos en este gif.
 
 
 
 
 
 
 
 

 
 
 
 
 

