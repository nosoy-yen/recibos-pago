# Política de privacidad — Recibos de Pago

Última actualización: 27 de octubre de 2026.

**Recibos de Pago** es un programa de código abierto para Windows que envía a
cada trabajador su recibo de pago por correo. Esta política explica qué hace
con tus datos.

## Qué permiso pide

Al iniciar sesión con Google, el programa pide un solo permiso:

- **Enviar correos** desde tu cuenta de Gmail (`gmail.send`).

**No pide ni puede** leer, buscar, borrar ni modificar tus correos, ni ver tus
contactos, tu bandeja de entrada o tus archivos.

## Qué datos usa y dónde se guardan

Todo se queda **en la computadora donde se usa el programa**, en archivos al
lado del programa:

| Archivo | Qué guarda |
|---|---|
| `sesion_gmail.json` | Tu correo y el permiso para seguir conectado |
| `empleados.csv` | La lista de nombres, cédulas y correos que tú escribes |
| `config.json` | Carpeta elegida, asunto y mensaje |
| `envios.log` | Un registro de a quién se envió cada correo |

Los recibos (PDF) se leen de la carpeta que tú eliges y solo se adjuntan a los
correos que tú mandas.

## Con quién se comparten

**Con nadie.** El programa no tiene servidor propio, no envía datos al autor ni
a terceros, no usa analítica ni publicidad. Los correos salen directamente de
tu computadora a Gmail (Google) y de ahí al destinatario que tú elegiste.

El uso que la aplicación haga de la información recibida de las API de Google
cumple con la
[Política de datos de usuario de los servicios de API de Google](https://developers.google.com/terms/api-services-user-data-policy),
incluidos los requisitos de uso limitado.

## Cómo borrar tus datos y quitar el permiso

- En el programa, **⚙ Configuración > Cerrar sesión** borra `sesion_gmail.json`
  y revoca el permiso en Google.
- Puedes revocarlo también en
  <https://myaccount.google.com/permissions>, eligiendo *Recibos de Pago*.
- Para borrar todo lo demás, elimina la carpeta del programa.

## Contacto

Dudas o problemas: abre un *issue* en
<https://github.com/nosoy-yen/recibos-pago/issues>.
