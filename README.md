# Pedidos Liberados SCS

## 1. Descripción

**Pedidos Liberados SCS** es una herramienta local para Windows destinada a automatizar el tratamiento y envío por correo electrónico de pedidos liberados generados desde SEFLOGIC/SAP para la **Gerencia de Servicios Sanitarios del Área de Salud de Lanzarote**.

El sistema revisa periódicamente una carpeta de red, identifica los documentos PDF correspondientes a pedidos, extrae del nombre del archivo los datos necesarios, consulta una base de proveedores mantenida en Excel, prepara el correo electrónico, adjunta el PDF y gestiona el archivo y la trazabilidad del proceso.

La aplicación está diseñada para trabajar de forma continua en un equipo Windows mediante el **Programador de tareas de Windows**, con una frecuencia prevista de **cada 5 minutos**.

La versión actual no necesita:

- Visual Studio.
- SDK de .NET.
- `dotnet`.
- Compilar ningún ejecutable.
- Microsoft Excel instalado para leer `Proveedores.xlsx`.
- Tener Thunderbird abierto.

El motor actual está implementado mediante **Windows PowerShell 5.1**.

---

## 2. Carpeta de trabajo

La carpeta principal vigilada por el sistema es:

```text
\\gerencialz.canariasalud\archivos\Logistica\COMPRAS\Pedidos Liberados
```

En esta carpeta se depositan los PDF pendientes de tratamiento.

El programa también utiliza o crea dentro de esta misma ubicación las siguientes carpetas:

```text
Pedidos Liberados
│
├── Proveedores.xlsx
├── pedidos enviados
│   └── AAAA
│       └── MM
│
├── reimpresiones
│
└── reportes
    ├── REPORT_PEDIDOS_AAAAMMDD.xlsx
    ├── INCIDENCIAS_PEDIDOS.xlsx
    └── panel_control.html
```

Las carpetas de año y mes se crean automáticamente cuando sean necesarias.

Ejemplo:

```text
pedidos enviados\2026\08
```

---

## 3. Formato esperado del PDF

El sistema espera documentos con una denominación equivalente a:

```text
NPERVAL_PRD0000001592_1_Proveedor_ 1000004856 Nº Pedido_ 4503204516.pdf
```

Del nombre del fichero obtiene los siguientes datos:

| Dato | Ejemplo |
|---|---|
| Usuario que libera/imprime | `NPERVAL` |
| Orden de impresión | `PRD0000001592` |
| Número de impresión | `1` |
| Código de proveedor | `1000004856` |
| Número de pedido | `4503204516` |

### Reglas actuales

- El número de proveedor puede tener distinta longitud.
- El número de pedido debe contener exactamente **10 dígitos**.
- El patrón debe contener:
  - `_Proveedor_`
  - `Nº Pedido_`
- El número situado antes de `_Proveedor_` determina si se trata de primera impresión o reimpresión.

---

## 4. Tratamiento de primeras impresiones

Cuando el número de impresión es:

```text
_1_
```

el documento se considera susceptible de envío automático.

El programa realiza, por orden, las siguientes operaciones:

1. Comprueba que el PDF ha terminado de copiarse.
2. Comprueba que el fichero no está bloqueado por otro proceso.
3. Analiza su nombre.
4. Extrae:
   - usuario;
   - PRD;
   - número de impresión;
   - proveedor;
   - pedido.
5. Calcula el hash SHA-256 del PDF.
6. Busca el proveedor en `Proveedores.xlsx`.
7. Obtiene:
   - nombre de empresa;
   - correo principal;
   - correos en copia.
8. Prepara el correo.
9. Adjunta el PDF.
10. Registra el estado `ENVIANDO`.
11. Intenta realizar el envío.
12. Si el servidor acepta el correo, registra el estado `ENVIADO_PENDIENTE_ARCHIVO`.
13. Mueve el PDF a:
    `pedidos enviados\AAAA\MM`.
14. Registra finalmente el estado `ARCHIVADO`.

---

## 5. Tratamiento de reimpresiones

Cuando el número de impresión es distinto de `1`, por ejemplo:

```text
_2_
_3_
_4_
```

el documento **no se envía automáticamente**.

Se mueve directamente a:

```text
reimpresiones
```

y queda registrado con el estado:

```text
REIMPRESION
```

La gestión posterior corresponde al personal de Logística.

---

## 6. Base de proveedores

La base se denomina:

```text
Proveedores.xlsx
```

y debe encontrarse en:

```text
\\gerencialz.canariasalud\archivos\Logistica\COMPRAS\Pedidos Liberados
```

El personal puede abrir y modificar directamente este Excel.

El programa lo vuelve a consultar durante cada ciclo, por lo que las altas o correcciones realizadas en el fichero se utilizan en posteriores ejecuciones sin necesidad de importar de nuevo la base.

### Columnas previstas

```text
COD_PROVEEDOR
NOMBRE_PROVEEDOR
EMAIL_1
EMAIL_2
ACTIVO
OBSERVACIONES
```

También pueden incorporarse posteriormente columnas:

```text
EMAIL_3
EMAIL_4
EMAIL_5
...
```

sin necesidad de modificar la lógica principal.

### Uso de los correos

- `EMAIL_1` = destinatario principal, campo **Para**.
- `EMAIL_2` y siguientes = destinatarios en **CC**.

### Campo ACTIVO

Los proveedores inactivos no se utilizan para el envío automático.

---

## 7. Proveedor no encontrado

Si el código de proveedor del PDF no existe en `Proveedores.xlsx`, el documento:

- no se envía;
- no se mueve a enviados;
- permanece pendiente;
- queda registrado con una incidencia;
- vuelve a ser comprobado en ciclos posteriores.

Estado utilizado:

```text
ESPERA_DATOS
```

Una vez que el personal añada el proveedor al Excel, el programa podrá procesarlo en una ejecución posterior.

---

## 8. Proveedor sin correo electrónico

Si el proveedor existe pero `EMAIL_1` está vacío:

- no se envía;
- el PDF permanece pendiente;
- se registra la incidencia;
- se vuelve a revisar en posteriores ciclos.

Estado:

```text
ESPERA_DATOS
```

Al informar posteriormente el correo en `Proveedores.xlsx`, el sistema podrá continuar automáticamente.

---

## 9. Correo electrónico

La cuenta remitente prevista es:

```text
logisticalz.scs@gobiernodecanarias.org
```

### Asunto

El asunto se construye de la siguiente forma:

```text
Pedido Gerencia de Servicios Sanitarios de Lanzarote nº 4503204516 – NOMBRE EMPRESA
```

### Cuerpo

El texto utilizado actualmente es:

```text
Adjunto remitimos Pedido nº: 4503204516 para NOMBRE DE LA EMPRESA

Rogamos confirmación de la recepción de este correo y/o activar en su cuenta de correo electrónico la confirmación automática de recepción de correos.

En caso de incidencias con el pedido contactar con: logisticalz.scs@gobiernodecanarias.org

Atentamente,
```

El PDF correspondiente al pedido se incorpora como adjunto.

---

## 10. Firma del correo

La herramienta no automatiza Thunderbird.

Al ejecutarse mediante el Programador de tareas de Windows, Thunderbird no necesita permanecer abierto y su firma automática no se incorpora por sí sola.

Existe el archivo:

```text
firma.html
```

en el que puede incorporarse la firma institucional que se desee añadir al final de los correos automáticos.

---

## 11. Confirmaciones de recepción y entrega

El sistema está preparado para solicitar, cuando el servidor Exchange y el servidor receptor lo permitan:

- confirmación de lectura;
- notificación de entrega;
- notificación de fallo o retraso.

Estas solicitudes dependen de la configuración de Exchange y del servidor de correo del destinatario.

Por tanto, no puede garantizarse que todos los proveedores devuelvan automáticamente una confirmación.

---

## 12. Exchange corporativo

El envío está preparado para realizarse directamente contra el servidor de correo corporativo.

La aplicación no necesita abrir Thunderbird.

Los parámetros se encuentran en:

```text
config.json
```

Los principales valores son:

```json
"Mail": {
  "Enabled": false,
  "SmtpHost": "PENDIENTE_CONFIGURAR",
  "SmtpPort": 587,
  "EnableSsl": true,
  "AuthenticationMode": "WindowsIntegrated",
  "FromAddress": "logisticalz.scs@gobiernodecanarias.org"
}
```

### Estado actual de seguridad

Por defecto:

```json
"Enabled": false
```

Esto significa que el programa puede probar:

- lectura de carpetas;
- reconocimiento de nombres;
- consulta de proveedores;
- clasificación;
- generación de estados;
- generación de informes;

sin enviar correos reales.

El envío sólo debe activarse cuando se hayan confirmado los parámetros reales del servidor Exchange.

---

## 13. Cuenta de servicio

La ejecución automática está pensada para realizarse mediante una **cuenta de servicio corporativa**.

La cuenta debe disponer, como mínimo, de:

- acceso de lectura a la carpeta principal;
- acceso de escritura;
- permiso para crear carpetas;
- permiso para mover PDF;
- acceso a `Proveedores.xlsx`;
- acceso a la carpeta de reportes;
- autorización para utilizar el sistema de correo configurado.

La tarea puede funcionar aunque no exista una sesión interactiva de usuario abierta.

---

## 14. Frecuencia de ejecución

La frecuencia prevista es:

```text
cada 5 minutos
```

La instalación mediante:

```text
INSTALAR_TAREA.cmd
```

crea una tarea en el Programador de tareas de Windows denominada:

```text
SCS - Pedidos Liberados
```

---

## 15. Protección frente a archivos todavía en copia

Antes de procesar un PDF, la herramienta comprueba:

- que han transcurrido unos segundos desde su última modificación;
- que puede abrirse para lectura;
- que ningún otro proceso mantiene un bloqueo exclusivo sobre el fichero.

Esto evita intentar enviar documentos que todavía se estén generando o copiando desde SEFLOGIC/SAP.

---

## 16. Control de duplicados

Para cada PDF se calcula:

```text
SHA-256
```

Este hash permite identificar documentos idénticos.

Aunque las reimpresiones se controlan principalmente mediante el número de impresión del nombre del fichero, el hash funciona como salvaguarda adicional para impedir envíos duplicados de un mismo PDF ya tratado.

---

## 17. Estados del expediente de envío

El programa mantiene un estado interno para cada documento.

### Estados principales

```text
PENDIENTE
ENVIANDO
ENVIADO_PENDIENTE_ARCHIVO
ARCHIVADO
```

### Otros estados

```text
REIMPRESION
ESPERA_DATOS
ERROR_TEMPORAL
INCIDENCIA_PERMANENTE
INCIDENCIA_NOMBRE
ERROR_ARCHIVO
REVISION_MANUAL
DUPLICADO
```

---

## 18. Protección frente a interrupciones durante el envío

Una de las principales medidas de seguridad funcional es el estado:

```text
ENVIANDO
```

Antes de realizar el envío, el documento queda registrado en este estado.

Si Windows, PowerShell o el equipo se interrumpieran durante esa operación y en la siguiente ejecución todavía apareciera como `ENVIANDO`, el sistema no vuelve a enviarlo automáticamente.

Lo convierte en:

```text
REVISION_MANUAL
```

para que una persona compruebe si Exchange llegó a aceptar o no el correo.

El objetivo es evitar un segundo envío accidental.

---

## 19. Correo enviado pero PDF no archivado

Si Exchange acepta el mensaje correctamente pero posteriormente se produce un error al mover el PDF, el estado queda como:

```text
ENVIADO_PENDIENTE_ARCHIVO
```

En posteriores ejecuciones:

- no se vuelve a enviar el correo;
- únicamente se intenta completar el movimiento del PDF.

Este control evita duplicar pedidos debido a un problema de archivos o de red posterior al envío.

---

## 20. Reintentos de envío

Para errores temporales de correo o red se utiliza actualmente:

1. Primer intento.
2. Segundo intento aproximadamente 5 minutos después.
3. Tercer intento aproximadamente 15 minutos después.
4. Si continúa fallando:
   `INCIDENCIA_PERMANENTE`.

El sistema deja entonces el documento para revisión.

---

## 21. Incidencias por nombre incorrecto

Si un PDF no cumple el patrón esperado, no se intenta interpretar de forma arbitraria.

Se registra como:

```text
INCIDENCIA_NOMBRE
```

El archivo queda pendiente de revisión.

---

## 22. Registro interno

El estado interno se guarda actualmente en:

```text
%ProgramData%\SCS\PedidosLiberados
```

Dentro de esta ubicación se almacenan:

- estado de los documentos;
- eventos;
- logs técnicos.

El estado no depende únicamente de que el PDF exista o no en la carpeta.

Esto permite reconstruir operaciones y evitar determinados reenvíos accidentales.

---

## 23. Logs

Se generan logs técnicos diarios con denominaciones similares a:

```text
PedidosLiberados_20260810.log
```

y se almacenan en:

```text
%ProgramData%\SCS\PedidosLiberados
```

Los logs permiten revisar errores de:

- acceso;
- lectura;
- Excel;
- archivos;
- correo;
- generación de reportes.

---

## 24. Histórico de eventos

El programa registra eventos asociados a cada documento, entre otros:

```text
ENVIANDO
ENVIADO
ARCHIVADO
REIMPRESION
ESPERA_DATOS
ERROR_ENVIO
ERROR_ARCHIVO
REVISION_MANUAL
DUPLICADO
```

Estos eventos alimentan los reportes diarios.

---

## 25. Report diario

Se genera:

```text
REPORT_PEDIDOS_AAAAMMDD.xlsx
```

Ejemplo:

```text
REPORT_PEDIDOS_20260810.xlsx
```

El report incluye campos como:

- fecha y hora;
- evento;
- pedido;
- proveedor;
- empresa;
- usuario;
- PRD;
- número de impresión;
- archivo;
- estado;
- intentos;
- error;
- destinatario principal;
- CC;
- hash.

---

## 26. Informe de incidencias

También se genera:

```text
INCIDENCIAS_PEDIDOS.xlsx
```

Este documento muestra las incidencias activas que requieren atención.

Entre otras:

- proveedor no encontrado;
- proveedor sin correo;
- nombre de fichero incorrecto;
- fallo permanente de correo;
- duplicado;
- envío indeterminado;
- error de archivo.

El Excel se vuelve a generar en los ciclos posteriores.

Si estuviera abierto y Windows impidiera actualizarlo, el programa podrá volver a intentarlo en una ejecución posterior.

---

## 27. Panel de control

El programa genera:

```text
panel_control.html
```

dentro de:

```text
reportes
```

Puede abrirse mediante:

```text
PANEL_CONTROL.cmd
```

El panel muestra actualmente:

- última ejecución;
- número de pendientes;
- enviados del día;
- incidencias;
- pedido;
- proveedor;
- empresa;
- estado;
- intentos;
- error;
- archivo.

No permite modificar `Proveedores.xlsx`.

La modificación de proveedores se realiza exclusivamente mediante Excel.

---

## 28. Ejecución manual

Puede forzarse una revisión mediante:

```text
EJECUTAR_AHORA.cmd
```

La versión actual valida previamente la sintaxis del motor PowerShell.

Si la validación es correcta, ejecuta un ciclo completo.

---

## 29. Validación de sintaxis

Antes de realizar pruebas puede ejecutarse:

```text
VALIDAR_SINTAXIS.cmd
```

Este archivo invoca:

```text
Validar-Sintaxis.ps1
```

y utiliza el parser del propio Windows PowerShell para comprobar que `PedidosLiberados.ps1` no contiene errores de sintaxis.

El resultado esperado es:

```text
[OK] Sintaxis PowerShell valida.
```

---

## 30. Instalación automática

Una vez validado el funcionamiento manual, puede instalarse la ejecución periódica mediante:

```text
INSTALAR_TAREA.cmd
```

Debe ejecutarse como administrador.

El instalador:

1. solicita la cuenta corporativa que ejecutará la tarea;
2. solicita sus credenciales;
3. registra la tarea;
4. configura la ejecución periódica cada 5 minutos.

---

## 31. Desinstalación

Puede eliminarse la tarea mediante:

```text
DESINSTALAR_TAREA.cmd
```

Debe ejecutarse con privilegios administrativos.

---

## 32. Configuración

La configuración principal está en:

```text
config.json
```

Incluye:

- carpeta raíz;
- nombre del Excel;
- carpeta de enviados;
- carpeta de reimpresiones;
- carpeta de reportes;
- frecuencia;
- tiempo de estabilización de archivos;
- número máximo de intentos;
- tiempos de reintento;
- directorio de estado;
- configuración SMTP/Exchange;
- remitente;
- confirmaciones de correo;
- firma.

---

## 33. Archivos principales del programa

```text
PedidosLiberados.ps1
```

Motor principal.

```text
config.json
```

Configuración.

```text
Proveedores.xlsx
```

Base de proveedores.

```text
firma.html
```

Firma opcional.

```text
VALIDAR_SINTAXIS.cmd
Validar-Sintaxis.ps1
```

Comprobación del código PowerShell.

```text
EJECUTAR_AHORA.cmd
```

Ejecución manual.

```text
INSTALAR_TAREA.cmd
Instalar-Tarea.ps1
```

Instalación de la tarea programada.

```text
PANEL_CONTROL.cmd
```

Acceso al panel.

```text
CONFIGURAR_CORREO.cmd
```

Abre la configuración de correo.

```text
DIAGNOSTICO.cmd
```

Diagnóstico básico.

```text
DESINSTALAR_TAREA.cmd
```

Elimina la tarea programada.

---

## 34. Flujo resumido

```text
PDF en carpeta raíz
        |
        v
Comprobar que está completo
        |
        v
Analizar nombre
        |
        +------------------------+
        |                        |
 impresión = 1             impresión > 1
        |                        |
        v                        v
 Buscar proveedor          REIMPRESIONES
        |
        +-----------------------------+
        |                             |
 proveedor correcto              falta dato
        |                             |
        v                             v
 Preparar correo                ESPERA_DATOS
        |
        v
 ENVIANDO
        |
        v
 Enviar por Exchange
        |
        +-----------------------------+
        |                             |
       OK                           ERROR
        |                             |
        v                             v
 ENVIADO_PENDIENTE_ARCHIVO      Reintentos
        |                             |
        v                             v
 Mover PDF                 INCIDENCIA_PERMANENTE
        |
        v
 pedidos enviados\AAAA\MM
        |
        v
 ARCHIVADO
```

---

## 35. Limitaciones actuales

La versión actual presenta deliberadamente las siguientes limitaciones:

1. El envío real está desactivado hasta configurar Exchange.
2. No existe edición de proveedores desde el panel.
3. No existe interfaz gráfica de escritorio.
4. El panel es un HTML generado con información de estado.
5. No se analiza el contenido interno del PDF.
6. Los datos del pedido se obtienen del nombre del fichero.
7. No se integra directamente con SEFLOGIC/SAP.
8. No se integra directamente con un gestor de expedientes.
9. La firma de Thunderbird no se obtiene automáticamente.
10. La confirmación de lectura depende del servidor y del destinatario.
11. Los pedidos en `REVISION_MANUAL` requieren intervención humana antes de cualquier posible reenvío.
12. Los errores permanentes no generan actualmente un correo interno de aviso; deben revisarse mediante `INCIDENCIAS_PEDIDOS.xlsx`.

---

## 36. Estado actual de implantación

La versión actual corresponde a una fase de **prueba funcional previa a la activación del envío real**.

El orden recomendado es:

1. Validar sintaxis.
2. Ejecutar manualmente con correo desactivado.
3. Comprobar reconocimiento de PDF.
4. Comprobar lectura de `Proveedores.xlsx`.
5. Comprobar reimpresiones.
6. Comprobar reportes.
7. Comprobar panel.
8. Obtener los parámetros corporativos de Exchange.
9. Realizar una prueba con una dirección controlada.
10. Activar el envío.
11. Instalar la tarea automática cada 5 minutos.
12. Validar el funcionamiento con la cuenta de servicio corporativa.

---

## 37. Principio de funcionamiento

El programa actúa únicamente como herramienta de automatización de la **remisión material de pedidos ya liberados**.

No:

- crea pedidos;
- modifica pedidos;
- aprueba pedidos;
- libera pedidos;
- cambia datos económicos;
- modifica SEFLOGIC/SAP;
- decide qué proveedor debe recibir un pedido.

Su función es identificar un pedido previamente generado, localizar los datos de contacto definidos por el personal, realizar el envío, archivar el documento y dejar evidencia del resultado.
