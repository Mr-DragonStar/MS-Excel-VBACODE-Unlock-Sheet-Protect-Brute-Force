# MS Excel VBACODE Unlock Sheet Protect Brute Force

Qué hace el código:
Intenta todas las combinaciones posibles de contraseñas formadas por:

Los primeros 11 caracteres son letras 'A' (ASCII 65) o 'B' (ASCII 66)
El 12º carácter prueba todos los caracteres imprimibles (ASCII 32 a 126)
Para cada combinación generada, intenta desproteger la hoja usando ActiveSheet.Unprotect
Si logra desprotegerla (ProtectContents = False), muestra un mensaje "Ok" y termina

Limitaciones:
Solo prueba contraseñas de exactamente 12 caracteres
Los primeros 11 caracteres solo son 'A' o 'B' (muy limitado)
Es extremadamente ineficiente para contraseñas largas o complejas
Puede tardar mucho tiempo (aunque en este caso el espacio de búsqueda es pequeño)

Consideraciones éticas:
Este tipo de código plantea cuestiones éticas y legales. Solo debería usarse:
En hojas propias donde hayas olvidado la contraseña
Nunca para acceder a archivos protegidos de otros sin autorización
