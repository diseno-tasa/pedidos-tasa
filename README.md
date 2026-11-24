# Pedidos de Material — Astillero TASA · Jefatura de Diseño

> App publicada en **https://diseno-tasa.github.io/pedidos-tasa/** (cuenta GitHub `diseno-tasa`).

Aplicación web (computadora y celular) para gestionar los pedidos de material de proveedores terceros por embarcación: grafo, Pos., reservas, plantillas, alertas y el Excel de carga.

Funciona **sin cuenta de Claude**:
- **GitHub Pages** publica la página (gratis).
- **Firebase** de Google guarda los datos y los usuarios (plan gratuito *Spark*).

---

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | La aplicación completa |
| `config.js` | Datos de **su** proyecto Firebase (se llena en el paso 4) |
| `firestore.rules` | Reglas de seguridad de la base de datos (se pegan en el paso 3) |
| `.nojekyll` | Archivo vacío que GitHub Pages necesita para publicar los archivos tal cual |

---

## Instalación paso a paso (≈ 20 minutos)

### 0. Descargue el respaldo de la app actual
Esto solo aplica si ya usaba la app en Claude.

1. Abra la app en Claude e ingrese como administradora.
2. Vaya a **Gestión → Eliminar datos → ⬇ Descargar respaldo (JSON)**.
3. Guarde el archivo en un lugar seguro: contiene las contraseñas.

### 1. Cree el proyecto en Firebase
1. Entre a <https://console.firebase.google.com> con una cuenta de Google (de preferencia una del área, no una personal).
2. Haga clic en **Crear un proyecto**, póngale un nombre (por ejemplo `pedidos-tasa`) y desactive Google Analytics, que no hace falta.
3. Quédese en el plan gratuito **Spark**.

### 2. Active el ingreso con usuario y contraseña
1. En el menú, vaya a **Compilación → Authentication** y haga clic en **Comenzar**.
2. En **Método de acceso**, elija **Correo electrónico/contraseña**, actívelo y guarde.

### 3. Cree la base de datos y sus reglas
1. En el menú, vaya a **Compilación → Firestore Database** y haga clic en **Crear base de datos**.
2. Ubicación: elija `southamerica-east1 (São Paulo)`, la más cercana a Perú.
3. Elija **Modo de producción**.
4. Vaya a la pestaña **Reglas**, borre lo que haya, pegue **todo** el contenido de `firestore.rules` y haga clic en **Publicar**.
   - Si la consola marca un error, copie el mensaje y consúltelo antes de seguir.

### 4. Copie la configuración en `config.js`
1. Haga clic en ⚙ **Configuración del proyecto**.
2. En **Tus apps**, haga clic en el ícono web **`</>`**, póngale un nombre a la app y regístrela. No necesita activar *Hosting*.
3. Firebase muestra un bloque `firebaseConfig = { apiKey: "...", ... }`. Copie esos valores dentro de `config.js`, entre las comillas de cada campo.

> Estos datos **no son secretos**: solo identifican su proyecto. Lo que protege la información son las reglas del paso 3.

### 5. Suba los archivos a GitHub
1. Cree una cuenta en <https://github.com> si no tiene una.
2. Haga clic en **New repository**, póngale un nombre (por ejemplo `pedidos-tasa`) y márquelo como **Public**.
   - Para publicar con GitHub Pages en una cuenta gratuita, el repositorio debe ser público.
   - El código queda visible para cualquiera, pero los datos no: solo entran los usuarios que usted cree.
3. Dentro del repositorio, haga clic en **Add file → Upload files** y arrastre `index.html`, `config.js`, `firestore.rules`, `README.md` y `.nojekyll`.
4. Haga clic en **Commit changes**.

### 6. Publique con GitHub Pages
1. En el repositorio, vaya a **Settings → Pages**.
2. En **Source**, elija **Deploy from a branch**; en **Branch**, elija `main` y `/ (root)`; luego **Save**.
3. Después de 1 o 2 minutos verá la dirección de la app: `https://SU-USUARIO.github.io/pedidos-tasa/`.

### 7. Autorice esa dirección en Firebase
1. En Firebase, vaya a **Authentication → Configuración → Dominios autorizados → Agregar dominio**.
2. Escriba `SU-USUARIO.github.io` (sin `https://` y sin la carpeta).

### 8. Configure la app por primera vez
1. Abra la dirección de la app. Verá la pantalla **Primera configuración**.
2. Cree su usuario administrador. Si usa el mismo usuario que tenía en Claude, la importación lo respeta.
3. Si tiene un respaldo, vaya a **Gestión → Eliminar datos → ⬆ Importar respaldo (JSON)** y elija el archivo del paso 0. La app vuelve a crear:
   - los usuarios, con sus mismas contraseñas;
   - el catálogo, las embarcaciones, las solicitudes, las plantillas, los periodos y las fotos.
4. Ingrese con un usuario de prueba de proveedor y confirme que solo ve lo suyo.

¡Listo! Comparta la dirección con su equipo y proveedores, junto con el usuario y la contraseña de cada uno.

---

## Uso diario

- **Proveedores:** no necesitan ninguna cuenta adicional. Abren la dirección e ingresan con el usuario y la contraseña que usted les dio.
- **Usuarios y contraseñas:** se gestionan en **Gestión → Usuarios**. Los administradores pueden ver la contraseña vigente de cada usuario, como se pidió.
- **Respaldos:** haga uno periódico desde **Gestión → Eliminar datos → ⬇ Descargar respaldo**.
- **Cambios en la app:** reemplace `index.html` en GitHub (**Add file → Upload files**) y los cambios se publican solos en 1 o 2 minutos.

## Límites del plan gratuito de Firebase (Spark)

| Recurso | Límite gratuito | Uso esperado |
|---|---|---|
| Lecturas | 50 000 por día | El catálogo completo se lee en solo 8 documentos, sin importar cuántos materiales tenga. Cada ingreso lee además las solicitudes visibles, las embarcaciones y los usuarios. |
| Escrituras | 20 000 por día | Holgado para el trabajo diario. |
| Almacenamiento | 1 GiB | Holgado. Las fotos referenciales se reducen a unos 50 KB cada una. |

Si un día se supera el límite, la app deja de responder hasta el día siguiente (hora del Pacífico de EE. UU.). Para evitarlo, elimine periodos antiguos ya cerrados, después de descargar un respaldo. Si algún día hiciera falta más capacidad, el plan *Blaze* cobra solo el exceso, a centavos por cada 100 000 operaciones.

## Notas de seguridad

- Los usuarios ingresan internamente como `usuario@pedidos-tasa.app`. Ese dominio no necesita existir. Si lo cambia en `config.js`, cámbielo también en `firestore.rules`, y no lo cambie después de crear usuarios.
- Las reglas impiden que un proveedor vea solicitudes de otras empresas, cambie su propio rol o lea contraseñas.
- La creación de cuentas está abierta en Firebase porque la app la usa para dar de alta usuarios. Una cuenta creada fuera de la app **no tiene acceso a nada**: el acceso depende de que exista su ficha en **Usuarios**.
- Pida a los usuarios que no reutilicen contraseñas personales, ya que los administradores pueden verlas.
