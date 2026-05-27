# proyecto_final
Generador de flujos de retrabajo para un proceso

Rosa Angela Loera Ornelas

Grupo: 942

Ya que comento que el trabajo podría realizarec en excel, opte por esta opción y Visual Basic porque me resultó más sencillo y fácil de manejar procesos con multiple pasos y subpasos.

#Proceso elegido

Elegí el ensamble de un mueble, ya que implica varios pasos principales y cada uno puede llegar a tener subprocesos.

Me base en el ejercicio de ejemplo.Tras leer el archico, decidí usar un enfoque basado en tablas para que sea mas facil de modificar en el futuro. Creé tres tablas:
- Proceso principal
- Flujos de trabajo
- Retrabajo por paso
En cada flujo de actividad el primer paso se identifica con el numero "1"

#Correccioens y desarrollo

Durante la escritura de la macro surgieron varios errores que fui corrigiendo con ayuda de DeepSeek (Gracias IA<3)

Los problemas principales que surgieron fueron:
- Al no tener activada la referencia Microsoft Scripting Runtime, no podía usar Dictionary.
  Entonces la solución fue usar CreateObject("Scripting.Dictionary")
- Cuando un flujo no tenía un paso con orden 1, el programa fallaba así que se agrego una validación y un mecanismo de fallback

#Resultado

Ejecuté la macro con el proceso de fabricacion de un mueble. Al verficar la salida, comprobé que si coincidia con los parámetros deseados.
Al final es muy fácil de configurar y modificar. Solo es necesari cambiar los valores en la hoja Config para adaptar sin tocar el código.
