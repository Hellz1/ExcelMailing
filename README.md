# ExcelMailing
Este proyecto utiliza la librería EPPlus para la generación de un Excel, con cifrado de contraseña y se envía vía SMTP usando Outlook. (haciendo uso de Stream, sin usar escritura en disco)

La idea de tener 2 proyectos separados es que sean independientes los aplicativos. (puedan vivir en servidores distintos). De éste modo, la generación de excel es independiente, y se puede trabajar el formato, sin afectar al proyecto que agrega contraseña.

Pueden ser modificados y mejorados ambos proyectos, pero son la base de la funcionalidad.
