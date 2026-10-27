# Recibos de Pago

Programa sencillo para Windows que le manda a cada trabajador **su** recibo de
pago por correo, sin tener que revisar los PDF uno por uno.

Los recibos llegan en un RAR con archivos nombrados por cédula
(`Reg_35-2026_T-12345678.pdf`). El programa guarda una lista con el nombre, la
cédula y el correo de cada persona, busca en la carpeta el PDF que lleve esa
cédula en el nombre, y con un botón manda todos los correos.

- Para enviar se inicia sesión con Gmail, como en cualquier página. Para
  Outlook, Yahoo u otros se usa una contraseña de aplicación.
- Si en la carpeta quedan recibos de meses anteriores, a cada persona se le
  envía solo el más nuevo.
- No hay que instalar librerías: solo Python, o ni eso con el `.exe`.
- Los datos quedan en archivos sencillos (`empleados.csv`, `config.json`) al
  lado del programa.
- También imprime los recibos marcados de un solo golpe.

## Qué necesita la computadora

Funciona en Windows 7, 10 y 11 (32 o 64 bits). Una PC con i5 y 8 GB de RAM va
de sobra. Hay dos formas de usarlo:

**A) Con el `.exe` (lo más fácil, no se instala nada).** Se compila una vez
—ver [Hacer el .exe](#hacer-el-exe-recomendado)— y en la computadora donde se
va a usar es solo copiar la carpeta y doble clic en `Recibos.exe`.

**B) Con Python.** Instalar **Python 3** desde
<https://www.python.org/downloads/windows/> marcando *Add python.exe to PATH*,
y abrir con `Recibos.bat`. No hace falta nada más: el programa usa solo lo que
viene incluido con Python.

Para que el programa abra los RAR solo, la computadora tiene que tener
**WinRAR** o **7-Zip**. Si no está ninguno, se abre el RAR a mano (clic derecho
> Extraer aquí) dentro de la misma carpeta y el programa igual funciona.

## Cómo se usa

La ventana tiene solo lo necesario: la carpeta a la izquierda, la lista de
personas en el medio, y los botones **Enviar** e **Imprimir** a la derecha. Lo
demás está en **⚙ Configuración**.

**Una sola vez:**

1. Abrir `Recibos.exe` (o `Recibos.bat` si se usa con Python).
2. **Elegir carpeta**: la carpeta donde se van a meter los RAR cada mes.
   Queda guardada.
3. **⚙ Configuración > Iniciar sesión con Gmail**: elegir la cuenta en el
   navegador y aceptar. Arriba queda *Se envía desde ...*. Si el correo no es
   Gmail, ahí mismo está *Usar otro correo* con
   [contraseña de aplicación](#contraseña-de-aplicación). En Configuración
   también se cambian el asunto y el mensaje (`{nombre}` y `{cedula}` se
   cambian por los de cada persona).
4. **+ Agregar persona** con nombre, cédula y correo, o en Configuración
   *Abrir lista en Excel* para pegar muchas de golpe. Doble clic en una
   persona para corregirla; clic derecho para quitarla.

**Cada mes:**

1. Copiar el RAR de los recibos en la carpeta. El programa se da cuenta solo
   en unos segundos y lo abre en una subcarpeta con su nombre y la fecha, así
   que el del mes siguiente se abre aunque se llame igual. Lo que hay en la
   carpeta se ve en el cuadro de la izquierda.
2. Cada persona tiene una casilla: **verde** si su recibo está, **roja** si
   no hay ningún PDF con su cédula. Si una persona tiene varios PDF (por
   ejemplo el de este mes y el del pasado), se usa el más nuevo según la fecha
   del archivo.
3. Marcar las casillas, o **Seleccionar todos** (marca solo a los que tienen
   recibo).
4. **Enviar**. En la columna *Estado* cada persona pasa a *➜ Enviando...* y
   luego a *✔ Enviado* (verde) o *✖ No se envió* (rojo), y debajo de los
   botones se ve la barra de avance. Los enviados se desmarcan solos, para no
   mandarles dos veces. Si alguno falla, al final sale un aviso con el motivo.
   Todo queda anotado en `envios.log`.
5. **Imprimir** manda los recibos marcados a la impresora que Windows tenga
   puesta por defecto.

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

## Iniciar sesión con Gmail

El botón usa el inicio de sesión oficial de Google (OAuth). El programa solo
pide permiso para **enviar** correos, nunca para leerlos, y no ve la
contraseña de la cuenta. La sesión queda guardada en `sesion_gmail.json` y
*Cerrar sesión* la borra y le quita el permiso al programa.

### Configurarlo (una sola vez, lo hace quien compila el programa)

Google exige que cada programa tenga su propio "proyecto". Es gratis y toma
unos 10 minutos:

1. Entrar a <https://console.cloud.google.com> con tu cuenta de Google.
   Arriba, *Seleccionar proyecto > Proyecto nuevo*, nombre `Recibos de Pago`,
   *Crear*.
2. Menú > *APIs y servicios > Biblioteca*, buscar **Gmail API** y darle a
   *Habilitar*.
3. Menú > *Google Auth Platform* (o *Pantalla de consentimiento de OAuth*) >
   *Comenzar*. Nombre de la app `Recibos de Pago`, tu correo de asistencia,
   público **Externo**, tu correo de contacto, aceptar y *Crear*.
4. En *Público*, darle a **Publicar app**. Si se deja en modo *Prueba*,
   Google cierra la sesión cada 7 días y hay que volver a entrar.
5. En *Clientes > Crear cliente*, tipo de aplicación **App de escritorio**,
   nombre `Recibos`, *Crear* y **Descargar JSON**.
6. Poner ese archivo (`client_secret_....json`) en la carpeta del proyecto, al
   lado de `recibos.py`. No hace falta cambiarle el nombre.
7. Doble clic en `hacer-exe.bat`. El `.exe` nuevo lleva ese archivo adentro y
   el botón ya funciona.

Al iniciar sesión, Google muestra **"Google no verificó esta app"**. Es
normal: la verificación es para apps con miles de usuarios. Se le da a
*Configuración avanzada > Ir a Recibos de Pago*. Así la pueden usar hasta 100
cuentas, de sobra para esto.

El `client_secret` no se sube a GitHub (está en el `.gitignore`): quien use
este código hace su propio proyecto de Google.

## Contraseña de aplicación

Para Outlook, Yahoo, un correo de empresa, o Gmail sin el botón. Estos
servicios ya no dejan que un programa entre con la contraseña normal. Hay que
generar una contraseña aparte, solo para este programa:

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

`empleados.csv`, `config.json`, `envios.log`, `sesion_gmail.json`,
`client_secret*.json` y los PDF **no** se suben a GitHub: ya están en el
`.gitignore`. Son nombres, cédulas, correos y accesos de gente real.

## Hacer el .exe (recomendado)

Es la forma de que en la computadora donde se va a usar **no haya que
instalar nada**, ni Python. Se compila una vez en cualquier PC con Python y
se copia el `.exe` por pendrive. Doble clic en **`hacer-exe.bat`**: instala
PyInstaller, y si encuentra el `client_secret` de Google lo mete dentro del
`.exe` para que funcione *Iniciar sesión con Gmail*.

Queda `dist\Recibos.exe`, un solo archivo de unos 11 MB. Al lado de él se
crean solos `empleados.csv`, `config.json` y `sesion_gmail.json`, así que
conviene dejarlo en una carpeta propia y no en el Escritorio suelto.

Dos cosas normales del `.exe`:

- **Tarda unos segundos en abrir.** Con `--onefile` se descomprime en memoria
  cada vez que arranca: medido, 4 segundos en una laptop actual, y más en una
  PC vieja con disco mecánico. Después de la primera vez va más rápido porque
  Windows lo deja en caché.
- **SmartScreen avisa "Windows protegió tu PC"** porque el `.exe` no está
  firmado. Se le da a *Más información > Ejecutar de todas formas*. Pasa una
  vez. Si el antivirus del trabajo lo bloquea del todo, no queda otra que
  instalar Python y usar `Recibos.bat`.

El `.exe` no se sube a GitHub (`dist/` está en el `.gitignore`): el repo es
para el código, el ejecutable se pasa por pendrive.

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

## Privacidad

El programa solo pide permiso para enviar correos y todo queda en la computadora
donde se usa. Ver [PRIVACY.md](PRIVACY.md).

## Licencia

MIT. Ver [LICENSE](LICENSE).
