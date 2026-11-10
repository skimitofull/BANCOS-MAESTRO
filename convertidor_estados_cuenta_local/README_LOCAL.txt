PROGRAMA LOCAL - ESTADOS DE CUENTA DIDACTICOS

INSTALACION
1. Ejecuta instalar_dependencias.bat una sola vez.
2. Ejecuta ejecutar_app.bat para abrir la aplicacion local.
3. El programa funciona en tu computadora y no envia archivos a internet.

FUNCIONES
- PDF a Excel: extrae movimientos y prepara un resumen.
- BANAMEX MICUENTA reconstruye fechas separadas, movimientos continuados entre
  paginas y las columnas de retiros, depositos y saldo. Conserva operaciones
  informativas sin importe y concilia cantidades, totales y saldo final contra
  el resumen impreso del documento.
- PDF a Excel ofrece dos modalidades para todos los bancos detectados:
  a) un movimiento por fila con el concepto completo en una celda;
  b) un renglon fisico del PDF por fila, con fecha e importes vacios en los
     renglones de continuacion para evitar cantidades duplicadas.
- Excel a PDF multibanco: genera Banorte, Scotiabank, Azteca, Banamex, BBVA,
  Santander, HSBC o el formato didactico generico.
- Generador didactico: crea ingresos y egresos ficticios con importes desiguales.
- El Generador didactico admite tres origenes: sin archivo base, un Excel base o
  un solo PDF base. Al usar PDF detecta el banco, conserva todos sus movimientos
  originales y agrega solamente las cantidades y montos nuevos solicitados.
- Con PDF base, las fechas del periodo provienen del estado de cuenta, el nuevo
  saldo inicial participa en la cadena diaria y el estilo de salida se selecciona
  automaticamente. También permite descargar el Excel para auditoria.
- En SALDO FINAL permite definir periodo, cantidades, monto total de ingresos,
  saldo inicial y saldo final deseado.
- En SALDO PROMEDIO MENSUAL solo exige saldo inicial, cantidades de ingresos y
  egresos y promedio deseado. Calcula ingresos, egresos y saldo final; de forma
  opcional acepta un ingreso mensual aproximado. El motor puede mover esa cifra,
  prueba calendarios alternativos y solo presenta una propuesta despues de
  validar promedio, liquidez, saldo minimo y saldo final. Esta modalidad usa el
  mismo motor con y sin archivo base, y también con Excel o PDF base.
- Si el promedio exacto no es viable, muestra la propuesta segura mas cercana,
  los limites inferior y superior estimados y, cuando aplica, el ingreso mensual
  requerido. Al aplicar una alternativa se fijan juntos el promedio y los
  ingresos totales que la hacen viable; asi la siguiente ejecucion no aumenta
  nuevamente la sugerencia. Los botones permiten copiar y recalcular.
- Permite conservar un Excel base y agregarle movimientos nuevos.
- Incluye periodos enlazados multibanco: recibe 2 o mas PDF, detecta cada
  banco, conserva sus movimientos, mezcla ingresos/egresos ficticios y arrastra
  el saldo final al periodo siguiente.
- En cada periodo se puede fijar una de tres metas: monto total de egresos
  nuevos, saldo final deseado o saldo promedio mensual deseado. En los dos
  ultimos casos el programa calcula automaticamente el monto de egresos.
- El saldo promedio mensual suma el saldo al cierre de cada dia natural y lo
  divide entre todos los dias del periodo mensual, incluidos dias sin movimientos.
- Al elegir saldo promedio mensual, la pantalla solo solicita el promedio
  deseado y las cantidades de ingresos/egresos. El motor calcula los montos
  sugeridos y el saldo final, buscando la combinacion viable con menor monto
  agregado. Saldo final y saldo minimo pueden fijarse en condiciones avanzadas.
- La propuesta muestra cantidades, totales sugeridos, promedio obtenido, saldo
  final y saldo minimo. El boton "Copiar propuesta a montos manuales" carga
  cantidades y montos en controles editables para aceptarlos o modificarlos.
- La opcion "Sugerir tambien cuantos movimientos nuevos agregar" propone las
  cantidades a partir de la densidad y proporcion de ingresos/egresos del PDF
  original. Puede desactivarse para escribir cantidades propias.
- Si el promedio exacto es matematicamente imposible con los movimientos
  originales y las restricciones elegidas, se conserva la alternativa segura
  mas cercana y se muestra su diferencia; nunca se aceptan saldos negativos.
- El modo "Fijar ingreso mensual y calcular egresos" permite indicar, por
  ejemplo, $1,250,000 de ingresos totales mensuales. Con todos los PDF cargados,
  la pantalla recalcula al cambiar un valor y muestra egresos compatibles,
  promedio obtenido, saldo final, saldo minimo y viabilidad.
- Cuando el ingreso fijado no alcanza, la guia informa el faltante, el ingreso
  total aproximado requerido y el promedio minimo viable con la distribucion
  actual, para decidir si ajustar el ingreso o el promedio.
- El plan de ingresos variables recibe un ingreso mensual promedio y un
  porcentaje maximo de variacion hacia arriba o abajo. Asigna un total distinto
  a cada periodo, conserva exactamente el promedio del conjunto y usa la semilla
  reproducible para obtener siempre el mismo plan hasta que el usuario la cambie.
- El orden de los importes no puede ser completamente ascendente ni descendente;
  se mezcla por semilla para alternar meses altos, bajos e intermedios.
- Los ingresos mensuales planeados incluyen centavos distintos de .00 para evitar
  importes artificialmente redondos, sin salir del porcentaje ni alterar el
  promedio exacto del conjunto.
- Cuando el plan de ingresos esta activo, tambien se respeta desde la modalidad
  de sugerencia automatica: el motor fija el ingreso mensual asignado y calcula
  egresos, saldo promedio y saldo final. El saldo inicial solo aporta liquidez y
  nunca se suma a INGRESOS_ORIGINALES, INGRESOS_FICTICIOS o INGRESOS_TOTALES.
- Ante una combinacion inviable, el asistente muestra opciones numericas: promedio
  maximo o minimo viable conservando ingresos, ingreso requerido conservando el
  promedio y nuevo promedio de ingresos del plan. Los botones permiten aplicar
  una alternativa o conservar los controles manuales para decidir otro valor.
- El mismo plan variable funciona al fijar SALDO FINAL: conserva el ingreso total
  asignado a cada mes y calcula los egresos requeridos. Si el saldo final es
  imposible, propone el saldo final maximo viable o el ingreso adicional necesario.
- SALDO FINAL ofrece dos modalidades: manual por periodo o saldo final promedio
  con variacion porcentual. El plan asigna saldos finales distintos, con centavos
  y orden no monotono; cada saldo final sigue pasando como apertura del siguiente
  periodo. Si una meta no es viable, tambien puede ajustar el promedio del plan.
- Los catalogos de conceptos de ingresos y egresos son independientes y editables;
  cada opcion se separa con punto y coma. Si no se modifican, se usan listas base
  con TRANSFER, TRANSFERENCIA, TRANSFERENCIA INT, COBRO, PAGO y DEPOSITO.
- En Scotiabank las transferencias recibidas y enviadas usan estructuras distintas:
  las recibidas incluyen variante, banco, fecha/hora, rastreo, nombre, cuenta y
  codigo; las enviadas usan SWEB, banco, beneficiario, cuenta/hora, rastreo y fecha
  de operacion. Los movimientos originales conservan sus renglones extraidos.
- La tabla Scotiabank usa tipografia de 7.2 puntos, separadores verticales de 0.72
  puntos, sombreado alternado y ninguna linea horizontal entre movimientos.
- El motor rechaza cualquier configuracion cuyo saldo final sea negativo.
- La regla contable es saldo inicial + ingresos - egresos = saldo final. Los
  egresos pueden superar los ingresos del mes cuando el saldo inicial cubre la
  diferencia, pero ningun saldo posterior a un movimiento puede ser negativo.
- Tambien valida la liquidez movimiento por movimiento. Puede adelantar un
  ingreso ficticio o posponer un egreso ficticio cuando sea necesario, pero
  nunca cambia la fecha de un movimiento original. Si aun asi no hay fondos,
  informa el faltante y el saldo inicial minimo requerido, y no genera archivos.
- La hoja Resumen muestra saldo inicial, flujo neto, parte del saldo inicial
  utilizada, saldo final calculado/reportado, diferencia contable, promedio
  diario, saldo minimo y el resultado de ambas reglas de validacion.
- El Excel maestro identifica cada renglon como ORIGINAL, FICTICIO o SISTEMA,
  conserva el SALDO_ORIGINAL y muestra el saldo recalculado de la simulacion.
- Permite descargar todos los PDF didacticos, un Excel por periodo y un libro
  maestro de auditoria.
- Las dos modalidades de Excel tambien estan disponibles en todos los periodos
  enlazados y bancos compatibles.
- Exporta el concepto completo en una fila o separado en varios renglones.
- Incluye salida PDF didactica Banorte, Scotiabank, Banco Azteca, Banamex,
  BBVA (dos perfiles), Santander, HSBC y formato generico.
- Las salidas bancarias priorizan las hojas de movimientos: columnas, bandas,
  renglones de continuacion, referencias y paginacion se ajustan por banco.
- Detecta y extrae movimientos de los bancos incluidos en las referencias.
- La semilla permite repetir exactamente el mismo ejercicio.
- Para saldo promedio mensual se puede elegir si el promedio debe quedar por
  encima, por debajo o cerca del saldo final. El calendario automático adelanta
  ingresos y retrasa egresos, o aplica la estrategia inversa, según el objetivo.
- Permite fijar tolerancia, concentración monetaria dentro de ventanas de días,
  máximo de movimientos nuevos por día, uso exclusivo de días hábiles y saldo
  diario mínimo. Las fechas de los movimientos originales nunca se modifican.
- La simulación en tiempo real muestra promedio mínimo y máximo estimados,
  relación promedio/final, diferencia contra el objetivo y propuesta numérica
  más cercana cuando una combinación no es viable.
- Se pueden fijar simultáneamente ingreso mensual, saldo promedio y saldo final.
  En ese caso el total de egresos queda determinado por la ecuación contable y
  el motor busca fechas compatibles con el promedio solicitado.

ARCHIVOS BASE
El Excel debe incluir cinco columnas: FECHA, DESCRIPCION, DEPOSITO, RETIRO y
SALDO. Tambien se reconocen varios nombres equivalentes de columnas.

DATOS FICTICIOS
Las identidades, RFC, cuentas, CLABE y referencias generadas son sinteticas.
Los PDF muestran la leyenda DOCUMENTO FICTICIO - USO DIDACTICO. No elimines
esta leyenda ni uses datos personales reales en materiales de clase.

ACTUALIZACION DE PLANTILLAS
El calculo de movimientos esta separado de las plantillas. Cuando cambie el
formato de un banco se actualiza solamente su modulo de presentacion/extraccion.
El extractor Scotiabank esta aislado en scotiabank_parser.py y no depende de la
interfaz, para facilitar su reutilizacion posterior en el proyecto PDF a Excel.
El enrutador universal esta aislado en universal_extractor.py e incluye la
decodificacion especial requerida por los PDF HSBC de referencia.

CALIBRACION Y REVISION
- Las medidas se obtienen directamente de los PDF originales: pagina, fuente,
  tamano, coordenadas, distancia entre lineas y altura de cada movimiento.
- Scotiabank usa alturas de 27.36 o 72.96 puntos segun el numero de renglones.
- BBVA 1 y BBVA 2 son perfiles independientes; no comparten tipografia,
  columnas ni separacion vertical.
- Banamex conserva cada renglon fisico y puede continuar un movimiento en la
  pagina siguiente.
- Para revisar una actualizacion se comparan el PDF original y el generado en
  imagen, y se validan automaticamente las coordenadas del texto.
- Scotiabank usa Arial Unicode MS Regular a 8.35 puntos, con interlineado de
  9.12 puntos, conforme al perfil EDO_CTA_MASTER_APP_V1_15.
- En Excel a PDF, Generador didactico y Periodos enlazados, Scotiabank ofrece
  la descarga de su grafica de movimientos como PNG con fondo transparente.
  Los importes, el saldo inicial y el saldo final se recalculan con los datos
  vigentes antes de generar cada imagen.
BANAMEX MICUENTA
----------------
Al agregar movimientos ficticios a este formato, el programa usa unicamente:

- PAGO RECIBIDO DE [BANCO] para ingresos.
- PAGO INTERBANCARIO A [BANCO] para egresos.

En Periodos enlazados > Bancos para transferencias BANAMEX MICUENTA se pueden
marcar los bancos permitidos. Los bancos desmarcados quedan excluidos de todos
los movimientos nuevos. Los conceptos que ya existen en el PDF se conservan.

Las hojas de continuacion MiCuenta admiten exactamente 50 renglones fisicos.
Cuando un movimiento rebasa el renglon 50, sus lineas restantes continuan en
la hoja siguiente, igual que en el documento de referencia. Los movimientos
usan BanamexDisplay Regular de 9 puntos y no llevan separadores horizontales.

Identificadores ficticios BANAMEX MICUENTA:
- Cuenta/CLABE: 18 digitos, prefijo real del banco seleccionado y digito
  verificador CLABE valido.
- Referencia: 7 digitos aleatorios y unicos dentro del ejercicio.
- Rastreo de PAGO INTERBANCARIO A: 18 digitos con patron 0859.
- Rastreo de PAGO RECIBIDO DE BBVA MEXICO: 30 caracteres con patron BMOVP.
- Rastreo de PAGO RECIBIDO DE HSBC: 10 caracteres con prefijo HSB.
- Para los demas bancos: 18 digitos con el codigo bancario correspondiente.
Ninguno de estos identificadores se genera como serie consecutiva. Los datos
originales del PDF no se sustituyen ni se reformatean.

Los movimientos ficticios MiCuenta no imprimen el concepto generico
TRANSFERENCIA INTERBANCARIA. PAGO RECIBIDO imprime primero todos sus datos,
termina con la clave de rastreo y conserva vacio su septimo renglon; deposito y
saldo se imprimen en ese renglon final. PAGO INTERBANCARIO muestra la referencia
sola y MISMO DIA, sin agregar un renglon vacio. Los originales se respetan.

Variedad BANAMEX MICUENTA:
- Los bancos autorizados se barajan y se reparten de forma equilibrada; primero
  se usan los que todavia no aparecian en el PDF base.
- El equilibrio se calcula por separado para ingresos y egresos, después de
  ordenar definitivamente los movimientos. Dos movimientos nuevos consecutivos
  nunca usan el mismo banco cuando hay al menos dos bancos compatibles.
- El programa extrae los nombres del PDF base y no los reutiliza en los
  movimientos nuevos. Tampoco repite nombres nuevos dentro del ejercicio.
- Si todos los bancos autorizados ya fueron usados, avisa y los rota de forma
  balanceada. Si se agotan los nombres simples, avisa y crea nombres compuestos.
- Las repeticiones que ya existan entre movimientos originales no se eliminan,
  porque el contenido del PDF base se conserva sin alteraciones.
