#Perl-procesamiento-masivo-Microsoft-Office-Primavera-P6-automatizacion-valor-agregado
Perl script para procesamiento masivo con archivo Microsoft Office Primavera P6 automatizacion valor agregado

##Introducción

La idea principal es describir una metodología para aumentar la automatización en la manipulación de archivos (Microsoft Office u otros) y comenzar a crear valor agregado. No describiré el Framework de cadena de valor de Michael Porter publicado el 2009 o de cómo aplicar gestión de procesos de negocios, sino ideas más simples que tienen un gran valor por su eficiencia. Los elementos claves para esta transformación son:

![Elementos claves transformación](/home/appweb7/repository.git/github/perl-presupuestos/aris-express/elementos-claves-transformacion.png  "Elementos claves en la transformación")


##Objetivos

Automatizar tareas repetitivas, evidenciar tareas, valor agregado.

##Definiciones

Proceso: Es un conjunto estructurado, coordinado de actividades automatizadas o no que se realizan en forma secuencial o paralela para lograr un propósito específico dentro de la organización. Un proceso es transversal, horizontal dentro del organigrama. Dentro de la gestión de procesos de negocios se postula: "Creando valor para el cliente, se crea valor para el negocio y creando valor para el negocio, se crea valor para los accionistas."

Valor Agregado: (Wikipedia) En términos económicos, el valor agregado es el valor económico adicional que adquieren los bienes y servicios al ser transformados durante el proceso productivo.
Razonamiento : Por qué automatizar ?

Todas las actividades manuales tienen un mayor grado de error, tardan mucho tiempo en lograr resultados, su producción es variable.

>las tareas repetitivas no agregan valor alguno a sus clientes (ni menos para el inversionista), la automatización abre un conjunto de escenarios que permiten una mayor innovación más allá de la frontera del negocio.

la entrega de servicios más completos y con una mejor experiencia para sus clientes (generan confianza basado en los resultados). Dicho de otro modo, todas aquellas actividades que sean repetitivas con una transformación conocida (fórmula, algoritmo predecible) serán automatizadas. Las variables que son mejoradas al automatizar permiten visualizar eficiencia en:


![Mejora procesos](/home/appweb7/repository.git/github/perl-presupuestos/aris-express/mejora-proceso.png  "Mejora procesos")

##Audiencia / Escenario aplicable (a modo de ejemplo)

Escenario aplicable es toda persona que dentro de sus actividades en forma sistemática realiza tareas repetitivas con la manipulación de archivos en cualquier formato. Supongamos a modo de ejemplo, que tenemos una oficina de estudio de propuestas en donde se desarrollan una gran cantidad de presupuestos en el área de minería y/o de construcción, en donde todo el análisis de precios unitarios se desarrollan en forma manual (previo a la carga de la metodología constructiva hacia el software Primavera P6 por ejemplo). El contexto del escenario aplicable es:


![Escenario aplicable de ejemplo](/home/appweb7/repository.git/github/perl-presupuestos/aris-express/alcance.png  "Escenario aplicable ejemplo")

##Comenzando el análisis

Supongamos que tenemos la siguiente planilla de Analisis de precios unitarios en formato Microsoft Excel, en la que deseamos extraer todos los recursos (mano de obra, equipos y herramientas, materiales, subcontratos, otros costos):

![Microsoft Office Excel Planilla Analisis precios Unitarios](/home/appweb7/repository.git/github/perl-presupuestos/imagenes/formato-planilla-ANAPU.png  "Microsoft Office Excel Planilla Analisis precios Unitarios")

En forma paralela, se desarrolla el documento de la metodología constructiva a la que posteriormente se le agregan todos los recursos calculados del análisis de precios unitarios. Si en la planilla Microsoft excel en forma de libros están por itemizado, se puede desarrollar un guión (script) por ejemplo en lenguaje perl que permita revisar esta misma planilla para obtener todos los recursos para la carga de la metodología constructiva en el software Oracle Primavera P6.

##Paradigma a romper

Cuando eliminas esas tareas repetitivas, comienza a crecer tu negocio en valor, comienzas a diseñar un resultado cuantificable, orientado hacia la satisfacción del cliente, comienzas en definitiva a entregar un mejor servicio. (Por la vía de la eficiencia económica y no por la vía de reducir costos a ciegas).


![Perl procesamiento masivo Microsoft Office Primavera P6 automatizacion Valor agregado](/home/appweb7/repository.git/github/perl-presupuestos/imagenes/perl-presupuestos.gif  "Perl procesamiento masivo Microsoft Office Primavera P6 automatizacion Valor agregado")

La animación anterior, corresponde a un ejemplo en lenguaje perl que permite recorrer una planilla en formato Microsoft Excel para obtener todos los recursos: mano de obra, materiales e insumos, equipos y maquinarias, herramientas y fungibles. A modo de ejemplo solamente, el script abre Microsoft Office Excel, recorre el documento (en la práctica, esto es oculto para darle velocidad). Estos listados de recursos dentro del proceso de trabajo, son entregados a recursos humanos (contratación de perfiles), adquisiciones (compra de materiales y arriendo de maquinaria), Programación para realizar la programación de la obra según la metodología constructiva y carga de todos los recursos obtenidos al Software Primavera P6.

Para desarrollar esta automatización, se pueden emplear cualquier lenguaje que existe en el mercado, pero en mi experiencia, los más adecuados son Python y Perl. Esto como una solución, junto con el levantamiento de requerimiento, el desarrollo y las pruebas toman por lo general entre dos o tres dias.

##Aplicación: Usos en otros formatos de archivos de Microsoft Office

Esta tipo de solución, se puede aplicar con archivos Microsoft Office Word como búsqueda de patrones tags, Planillas contables para el procesamiento y validación y cálculo de información en cualquier tipo de situación en donde la transformación (algoritmo) sea conocido.

##Valor Agregado

Dónde comienza realmente percibir el valor agregado el cliente ? o mejor dicho dónde comienza a llamar la atención la propuesta de valor. La respuesta está en el conjunto de atributos que el productos o servicios logra entregar en forma explícita hacia el cliente como una experiencia excepcional, es decir satisfacer aquellos elementos que el cliente no percibía o no existe una propuesta similar (lo nuevo), mejoras en el desempeño (nivel de atención eficiente, mejoras en la tecnología), también la customización, el adaptar el producto a lo que el cliente necesita, hacer que el trabajo se realice (desarrollar trabajos en los que nadie quiere trabajar), mejoras en el diseño ya sea estético o funcional, diferenciación a través de marcas/status, precios, reducción de costos, reducción de riesgos, accesibilidad.