# Recibos de Pago

Programa sencillo para Windows que le manda a cada trabajador **su** recibo de
pago por correo, sin tener que revisar los PDF uno por uno.

Los recibos llegan en un RAR con archivos nombrados por cédula
(`Reg_35-2026_T-12345678.pdf`). El programa guarda una lista con el nombre, la
cédula y el correo de cada persona, busca en la carpeta el PDF que lleve esa
cédula en el nombre, y con un botón manda todos los correos.

- No hay que instalar librerías: solo Python.
- Los datos quedan en archivos sencillos (`empleados.csv`, `config.json`) al
  lado del programa.
- También imprime los recibos marcados de un solo golpe.

## Qué necesita la computadora

Funciona en Windows 7, 10 y 11 (32 o 64 bits). Una PC con i5 y 8 GB de RAM va
de sobra.

1. Instalar **Python 3** desde <https://www.python.org/downloads/windows/>
   (en el instalador marcar *Add python.exe to PATH*).
2. Nada más. El programa usa solo lo que viene incluido con Python.

Para extraer los RAR desde el programa hace falta tener **WinRAR** o **7-Zip**
instalado. Si no está ninguno, se extrae el RAR a mano (clic derecho > Extraer
aquí) y listo, el programa igual funciona.

## Cómo se usa

1. Doble clic en **`Recibos.bat`** para abrir el programa.
2. **Carpeta con los recibos**: dar a *Importar RAR...* y elegir el RAR que se
   descargó, o dar a *Elegir carpeta...* si ya se extrajo a mano. Esa carpeta
   queda guardada; la próxima vez solo se mete el RAR nuevo.
3. **Correo de quien envía**: poner el correo y la *contraseña de aplicación*
   (ver abajo). Dar a *Probar conexión* para confirmar que entra.
4. **Trabajadores**: cargar la lista una sola vez, con *Agregar* o con
   *Abrir lista en Excel* para pegar muchos de golpe. Queda guardada.
5. La columna **Recibo** dice quién tiene su PDF (`listo`) y quién no
   (`sin recibo`, en rojo).
6. Marcar con un clic en la primera columna, o dar a *Marcar todos* /
   *Solo los que tienen recibo*.
7. **Enviar correos a los marcados**. En el *Registro* de abajo se ve cuál
   salió bien y cuál falló. Ese registro también queda en `envios.log`.
8. **Imprimir los marcados** manda los PDF a la impresora que Windows tenga
   puesta por defecto.

El asunto y el mensaje se pueden cambiar. Si se escribe `{nombre}` o
`{cedula}`, el programa los reemplaza por los de cada persona.

### La lista de trabajadores

Es el archivo `empleados.csv`, con tres columnas:

```csv
nombre,cedula,correo
Juan Perez,12345678,juan.perez@correo.com
```

En `empleados-ejemplo.csv` está el formato. Para empezar, copiar ese archivo
como `empleados.csv` y cambiar los datos, o cargarlos desde el programa.

La cédula se compara solo por los números, así que da igual escribirla
`12.345.678`, `V-12345678` o `12345678`.

## Contraseña de aplicación

Gmail, Outlook y Yahoo ya no dejan que un programa entre con la contraseña
normal. Hay que generar una contraseña aparte, solo para este programa:

- **Gmail**: activar la verificación en dos pasos y luego crear una contraseña
  de aplicación en <https://myaccount.google.com/apppasswords>. Son 16 letras.
- **Outlook / Hotmail**: *Seguridad > Opciones de seguridad avanzadas >
  Contraseñas de aplicación*.

Esa contraseña se pega en el campo *Contraseña de aplicación*. Si se marca
*Recordar la contraseña*, queda guardada en `config.json` en esa computadora
(texto plano, así que `config.json` nunca se sube a GitHub). Si no se marca,
hay que escribirla cada vez que se abre el programa.

> Si el correo del trabajo es de la empresa (no Gmail), hay que pedirle a
> informática el **servidor SMTP** y el **puerto**, elegir
> *Otro (escribir a mano)* y ponerlos ahí.

## Cuidado con los datos

`empleados.csv`, `config.json`, `envios.log` y los PDF **no** se suben a
GitHub: ya están en el `.gitignore`. Son nombres, cédulas, correos y
contraseñas de gente real.

## Hacer un .exe (opcional)

Si en la computadora del trabajo no se puede instalar Python, se arma un
ejecutable en cualquier PC con Python y se copia la carpeta:

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name Recibos recibos.py
```

El `.exe` queda en `dist\`. Al lado de él se crean solos `empleados.csv` y
`config.json`.

## ¿Y por WhatsApp?

No está hecho, a propósito. Para mandar mensajes automáticos por WhatsApp hay
dos caminos y ninguno sirve bien para esto:

- Automatizar WhatsApp Web (con el teléfono) **va contra los términos de uso** y
  por mandar muchos mensajes seguidos a contactos que no escribieron primero,
  WhatsApp bloquea el número. Justo lo que no se quiere.
- La vía legal es la **WhatsApp Cloud API** de Meta: hace falta una cuenta de
  WhatsApp Business, un número aparte, verificar la empresa, y cada mensaje
  tiene que usar una *plantilla* aprobada por Meta y se paga por conversación.

Si algún día se quiere, lo correcto es la Cloud API. Por ahora el correo hace
el trabajo y no tiene riesgo de bloqueo.

## Licencia

MIT. Ver [LICENSE](LICENSE).
