# word2excel

Este repositorio fue creado por motivo de la siguiente [pregunta][1] en StackOverflow en Español:
----
> Gente, espero que entiendan que esto es una situacion especial. Es una pregunta que no cumple con las reglas, pero ahora mismo el mundo es un desastre y realmente necesito si alguien puede darme una mano con esto. Llenenme de negativos si quieren pero por favor, no cierren la pregunta.

Trabajo en el Ministerio de Seguridad de la Provincia de Buenos Aires y estamos cargando a mano toda denuncia referida al COVID-19. El tema es que este virus está creciendo exponencialmente, por lo que nuestro trabajo tambien lo hace, y estamos atrasados con la estadística, en parte por eso la cifra que se conoce sobre los casos es mas baja de lo que en realidad es. Sin mas preambulo, les paso a mostrar lo que tengo:

Nos llega un archivo .doc (pregunte en el chat antes de realizar la pregunta acá, y me preguntaron la version de Word. En realidad trabajamos con distintas versiones, yo en casa tengo la ultima version porque soy estudiante y MS la regala y actualiza siempre, pero en el trabajo tengo el 2010, creo que puede ser indiferente la version ya que yo quiero trabajar con texto, podria inclusive copiarlo y llevarlo a un notepad por ejemplo)

Este archivo .doc tiene unos 100 textos que tienen mas o menos esta forma

> Averig.Ilicito – Pto. Ilícito (COVID-19 Coronavirus) – Femenina Mayor
> **19/03** – PU.**81027** – Alta **07:06**hs – CP Campana; 05:00hs. Se
> recepcionó llamado a XXXX (Dda. XXXX 1000), dando cuenta que su vecina
> XXX, víspera regreso proveniente de Brasil, no cumpliendo con el
> protocolo de aislamiento.  Personal procedió a brindarle los
> correspondientes números telefónicos para tales casos, dando aviso a
> personal de Salud.

Cada uno de esos textos deben ser pasados a tabla en una base de datos, que tiene unos 20 campos. Yo quisiera al menos poder sacar los 3 datos que puse en negrita, que son, la fecha, el numero de Parte Urgenta, que esta siempre despues de ``PU.`` y la hora, que siempre es la que esta despues de la palabra ``alta`` (ya que puede haber otras horas en el parrafo, de todas maneras la hora que busco es la primera que aparece, asi que pueden tenerse esos dos criterios, o la primer aparicion, o despues de "alta")

**Si no sabes como hacerlo pero conoces herramientas que puedan hacerlo, tambien me seria de gran ayuda**

La cosa es que estamos pasandolos a mano, y este virus crece a pasos agigantados con ello las denuncias, y no estamos pudiendo cargar esto.

La idea es que cada nuevo texto (podria identificarse por la aparicion de la cadena ``PU.``) cree una nueva fila (puede ser en una tabla de excel o de access, da igual) y coloque esos 3 campos. 

Espero sepan entender una pregunta excepcional en un momento excepcional.

[1](https://es.stackoverflow.com/questions/338584/automatizar-carga-de-documento-pasar-de-word-a-excel-o-similar/338607#338607)