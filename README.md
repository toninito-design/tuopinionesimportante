# tuopinionesimportante
## Cloudflare Workers

La configuración de Wrangler compila la aplicación con `python3 worker/build.py`.
Workers Builds debe usar la raíz del repositorio y el comando de despliegue
`npx wrangler deploy`.

Crea un bucket **privado** R2 llamado `tuopinionesimportante-historico` en la misma
cuenta. Configura un dominio del Worker en Cloudflare y una aplicación de
Cloudflare Access para todo ese dominio con una política que permita únicamente
a las personas autorizadas. Define `ACCESS_ISSUER` como la URL del equipo de
Access (`https://equipo.cloudflareaccess.com`, sin barra final) y `ACCESS_AUD`
como el identificador de audiencia de esa aplicación. Configura las variables
en el entorno de ejecución del Worker. La configuración deshabilita
`workers.dev` y las URL de vista previa. El Worker rechaza solicitudes que no
presenten una identidad de Access válida para la aplicación.

Este despliegue comienza con su propio histórico. La copia del repositorio
contiene el código y no importa los datos almacenados en el sitio original.
