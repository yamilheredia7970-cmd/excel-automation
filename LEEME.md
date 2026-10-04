# Carga de Ingresos Brutos (AGIP / e-Sicol) en el Excel

Entra a AGIP con el CUIT de cada cliente, abre la DDJJ de cada mes (si tiene rectificativas, **solo la última**), lee los conceptos de la hoja DETALLE y los carga en la planilla. Primero todo 2024 y después todo 2025. Retoma siempre desde lo que falta.

## Uso diario
Con el Excel **cerrado** (si se edita mientras corre, el script lo pisa al guardar):

```
python iibb_agip_scraper.py --excel "RUTA\IIBB ANUAL 2025-4.xlsx" --sin-headless --pausa-accion 0.5
```

- Se puede cortar con Ctrl+C y volver a correr: sigue donde quedó.
- Cada ejecución crea su `iibb_agip_FECHA.log`. Al final trae el resumen y las listas: *REVISAR A MANO*, *Meses sin DDJJ presentada*, *Errores*, *No se llegaron a procesar*.
- Si algo falla se guarda una captura y el HTML en `capturas_errores/` (no se suben a git: tienen datos de clientes).
- Si AGIP o internet no responden, espera 5 minutos y reintenta (hasta ~1 hora) antes de cortar.

## Otras opciones (solo si hacen falta)
| Opción | Para qué |
|---|---|
| `--auditar` | Solo lectura. Lista los meses cuyas cuentas no cierran (anticipo − retenciones − percepciones − créditos ≠ importe a pagar / saldo a favor). Suele faltar una actividad o un crédito. |
| `--corregir` | Relee de AGIP **solo** esos meses y reemplaza lo que difiera. Se combina con `--anios` y `--cuit`. |
| `--rehacer` | Relee los meses indicados (con `--cuit` / `--meses`) aunque ya tengan datos. AGIP manda. |
| `--cuit A,B` `--meses 1,2` `--anios 2025` | Filtros para probar o corregir. |
| `--restaurar-colores` | Recupera del backup (`*.backup.xlsx`) los colores puestos a mano. |
| `--dry-run` | Solo lectura: resumen de clientes y meses pendientes. |

## Reglas acordadas
- **Colores de "Lista de Clientes - IIBB"**: verde = sin meses pendientes; dorado = pendiente sin error (incluye meses sin DDJJ presentada); rojo = hubo error en la corrida. El azul (DDJJ HECHA) y cualquier color puesto a mano **nunca se tocan**.
- **"No se puede entrar"** se escribe junto al CUIT (columna C en 2024, F en 2025) cuando ninguna contraseña funciona o el cliente no tiene CUIT.
- **Contraseña**: la escrita junto al CUIT en la hoja del cliente; si no hay o falla, las dos por defecto del script.
- **Otros créditos** = lo que AGIP aplicó como saldo a favor del período anterior; si AGIP no muestra nada, el saldo a favor del mes anterior de la planilla (enero usa diciembre del año anterior).
- **Ceros**: retenciones, retenciones bancarias, percepciones, impuestos internos, pago a cuenta, otros créditos y saldo a favor quedan en blanco si AGIP dice 0. Base, anticipo, importe a pagar y total pagado se escriben aunque sean 0. Intereses = total pagado − importe a pagar (fórmula).
- **Varias actividades**: cada una en su propio bloque (base, anticipo, alícuota, código). En las hojas `DIA N - 2025` el script crea el bloque de una actividad nueva; en las de 2024 (armadas a mano) **no agrega filas**: lo que falte queda en *REVISAR A MANO* para cargarlo a mano.
- **Qué pisa**: completa celdas vacías; un 0 de relleno y la alícuota se reemplazan; cualquier otro número distinto no se pisa y queda en *REVISAR A MANO*. Con `--rehacer` / `--corregir` manda AGIP, salvo que **un 0 de AGIP nunca borra** una retención, percepción, crédito o pago a cuenta ya cargado (queda en *REVISAR A MANO*).
- **Lecturas dudosas**: cada mes se verifica al leerlo (que no falten datos y que la cuenta cierre: anticipo − retenciones − percepciones − créditos = importe a pagar o saldo a favor). Si no cierra se vuelve a leer, hasta 3 veces y esperando más cada vez. Si sigue sin cerrar se carga lo leído y queda en *REVISAR A MANO* como "lectura dudosa". Los días en que AGIP anda lento la corrida tarda más.
- **Clientes repetidos** en la lista: se carga el primer bloque (conviene borrar el repetido). Clientes de la lista sin bloque: se les crea uno al final de la hoja de su DIA.

## Limitaciones conocidas
- AGIP no tiene API: se maneja la pantalla de e-Sicol (ventanas flotantes), así que un cambio del sitio puede romper un paso. El log y las capturas muestran en cuál.
- Clientes que representan a otras personas: se elige al propio cliente en "Seleccione un representado".
- e-Sicol a veces muestra "Error interno": el script acepta el cuadro y reintenta la sección; si insiste, ese dato queda sin cargar y avisado.
- Si e-Sicol responde lento, las ventanas pueden abrir vacías (con $0,00 por defecto). Por eso se verifica cada mes y se relee (ver *Lecturas dudosas*); sin esa verificación aparecían retenciones y percepciones en blanco, o la liquidación sin cargar.
- Se probó contra copias simuladas de e-Sicol y contra los logs y el Excel reales; AGIP no es accesible desde donde se desarrolló.
