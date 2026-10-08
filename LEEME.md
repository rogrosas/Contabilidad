# Contabilidad del negocio

Programa para registrar **ventas, compras, mermas y gastos**, controlar el
**stock** y la **caja diaria**, y sacar **planillas Excel por mes o por rango de
fechas**. Es muy liviano: abre en un par de segundos y usa unos 50 MB de RAM,
así que anda bien en equipos modestos (Core i3, 4 GB). No necesita internet.

## Cómo abrirlo

Haga doble clic en **`Iniciar Contabilidad.bat`** (o en `contabilidad.pyw`).

Requisito: Python 3.10 o más nuevo (https://www.python.org/downloads/). Al
instalarlo, marque **"Add python.exe to PATH"**. No hay que instalar nada más.

## Cómo se usa

1. **Productos**: cargue sus productos con su costo (precio de compra), su
   precio de venta y, si quiere aviso de reposición, un **stock mínimo**.
2. **Ventas**: cada venta es un ticket con uno o varios productos.
   - Escriba parte del nombre → **Enter** → cantidad → **Enter**. Repita.
   - Al terminar: **Enter** con el campo vacío, escriba con cuánto paga el
     cliente (muestra el vuelto) y **Enter** o **F12** para guardar.
   - Elija el medio de pago: efectivo, débito, crédito, transferencia u otro.
   - Doble clic en una venta registrada para abrirla y corregirla.
3. **Compras**: igual que las ventas. El costo de cada producto se actualiza
   solo con el precio de la compra más reciente.
4. **Mermas**: mercadería perdida (vencida, rota, robada). Se valoriza al costo.
5. **Gastos**: alquiler, luz, sueldos, impuestos… todo lo que no es mercadería.
6. **Caja**: al final del día cuente el efectivo. El programa calcula cuánto
   debería haber y le dice si **cuadra, falta o sobra**.
7. **Reportes y Excel**: elija **Mes**, **Día** o **Rango de fechas** y
   presione **Exportar a Excel**. Los archivos quedan en la carpeta `reportes`.

### IVA (Chile, 19%)

- **Todos los precios se cargan CON IVA incluido**, como en la góndola. El
  programa separa el neto y el IVA.
- En **Productos**, marque como **Exento** lo que no lleva IVA. Por defecto
  todo es "Afecto 19%".
- En **Compras** y **Gastos**, marque **"Con factura"** cuando tenga factura:
  así ese IVA se descuenta como **IVA crédito**. Sueldos, alquiler e impuestos
  vienen desmarcados; el programa recuerda lo que usó la última vez en cada categoría.
- **IVA a pagar** = IVA débito (ventas) − IVA crédito (compras y gastos con
  factura). Si un mes queda a favor, ese **remanente** se descuenta el mes
  siguiente, como en el **Formulario 29**.
- La tasa se cambia en *Archivo → Tasa de IVA*. Lo ya registrado conserva su tasa.

> El cálculo de IVA es referencial: no incluye el reajuste UTM del remanente,
> PPM ni otros impuestos. Verifique con su contador antes de declarar.

### Qué significa cada número

- **Ventas netas** = ventas sin IVA (el IVA no es plata del negocio).
- **Ganancia bruta** = ventas netas − lo que costaron esos productos (sin IVA).
- **Resultado neto** = ganancia bruta − mermas − gastos. *Lo que realmente ganó.*
- **Flujo de caja** = ventas − compras − gastos, con IVA. *Dinero que entró menos el que salió.*
- Montos en pesos sin decimales. Al escribir: `1.500` es mil quinientos y
  `1,5` es uno y medio (útil para kilos).

> Si vende un producto que todavía no tiene costo cargado, el programa avisa
> en Reportes. Cuando cargue el costo (o registre la compra), esas ventas se
> corrigen solas.

### Qué trae la planilla Excel

| Hoja | Contenido |
|---|---|
| Resumen | Resultado sin IVA, IVA del período, dinero, medios de pago y gastos por categoría |
| Por mes | Una fila por mes (si el período abarca varios meses), con IVA |
| IVA mensual (F29) | Ventas afectas y exentas, IVA débito, crédito, remanente e IVA a pagar |
| Por producto / Por categoría | Vendido, comprado, mermas, IVA, ganancia neta y margen |
| Ventas / Compras / Mermas / Gastos | Detalle de cada registro con neto e IVA, con filtros |
| Cierres de caja | Los cierres del período con sus diferencias |
| Inventario actual | Stock, valor y qué productos hay que reponer |

### Stock

- Abajo a la derecha aparece **"⚠ N productos para reponer"**. Haga clic para verlos.
- **Ajustar stock (conteo)**: en Productos, cuente la mercadería y escriba lo
  que hay. Si falta, se registra como merma; si sobra, se suma.
- **Ver historial** (o doble clic en un producto): todas sus entradas y salidas.

## Sus datos y respaldos

- Todo se guarda en `datos\negocio.db`. Para pasar el programa a otra
  computadora, copie la carpeta completa.
- **Respaldo automático**: cada día, al abrir el programa (se guardan los últimos 30).
- **Archivo → Guardar respaldo en pendrive**: haga esto al menos una vez por
  semana; si la computadora se rompe, los datos quedan a salvo. El programa le
  recuerda si pasó más de una semana.
- **Archivo → Restaurar un respaldo**: vuelve a los datos de una copia. Antes
  de restaurar, guarda automáticamente los datos actuales por si acaso.
- Si algo falla, el programa avisa y guarda el detalle en `datos\errores.log`.

## Atajos

| Tecla | Acción |
|---|---|
| Ctrl+1 … Ctrl+7 | Cambiar de pestaña |
| F12 | Cobrar / guardar el ticket |
| Enter | Pasar al campo siguiente / agregar / guardar |
| Supr | Quitar línea o eliminar registro seleccionado |
| Doble clic | Abrir o corregir un registro |
| ↑ / ↓ en una fecha | Sumar o restar un día |
| F4 en una fecha | Abrir el calendario |
| F5 | Actualizar |
