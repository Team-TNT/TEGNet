# TEGNet

Código fuente y documentación de soporte del proyecto TEGNet (www.tegnet.com.ar). El proyecto es un conjunto de aplicaciones desarrolladas con el objetivo de posibilitar a los usuarios disfrutar de un popular juego de mesa en forma remota y/o con jugadores virtuales.

## Cómo empezar

El propósito de este repositorio es proporcionar acceso al código fuente del proyecto para aquellos que tengan curiosidad por saber como funciona. Para quienes deseen correr la aplicación en sus propios entornos, iremos agregando información conforme tengamos tiempo para hacerlo.

Como el proyecto fué desarrollado hace unos cuentos años (2001), las herramientas requeridas para su ejecución ya quedaron hace tiempo obsoletas, haciendo que solo compilar el proyecto sea un gran desafío.

El proyecto consta de 3 componentes fundamentales:
1. El [Servidor](Servidor/), que maneja las partidas y la comunicación entre los distintos jugadores (locales y remotos). Se requiere una instancia de Servidor corriendo por cada partida (cada instancia solo puede procesar una partida simultanea).
1. El [Cliente](TegNet/), que contiene la UI de la aplicación y se conecta con un Servidor para ejecutar una partida. Permite la conexión con una partida existente, o el inicio de una nueva, así como también la configuración del juego y la ejecución propiamente dicha.
1. El [Jugador Virtual](Jv/), contiene la lógica requerida para generar uno o mas bots que emulan jugadores en una partida.

Para mas información acerca de la arquitectura del proyecto y sus componentes vea las [Especificaciones Técnicas (o Manual de Sistema, como le decíamos entonces)](Docs/technical-specs.md).

### Prerequisitos

Como hace mucho tiempo que no corremos el proyecto no estamos seguros de qué componentes exáctamente se requieren para correr la aplicación. Una vez que logremos volverlo a correr iremos completando con la información necesaria.

Las aplicaciones fueron desarrolladas en [Visual Basic 6](https://msdn.microsoft.com/en-us/vstudio/ms788232) con lo que, cómo mínimo, se requiere una versión de Visual Studio que permita buildear esa versión.

Los datos de las partidas y las misiones, etc. se persisten en una base de datos Access.

Para la conexión con la base de datos se requiere Microsoft Data Access Component (MDAC)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Ariel Clochiatti**
* **Emiliano Cavia**
* **Guido Pons** - [guidopons](https://github.com/orgs/Team-TNT/people/guidopons)
* **Guillermo Giannotti** - [guillegiannotti](https://github.com/orgs/Team-TNT/people/guillegiannotti)
* **Javier Rebagliatti** - [jrebagliatti](https://github.com/orgs/Team-TNT/people/jrebagliatti)

See also the list of [contributors](https://github.com/Team-TNT/TEGNet/contributors) who participated in this project.

## License

TBC

## Acknowledgments

* TBC
