# Guía de uso — PMP Generator

Esta herramienta arma sola el **PMP de la semana siguiente**: actualiza el
Control, la Matriz de disponibilidad y, si quieres, genera un cuadro resumen para
compartir. No necesitas instalar nada ni saber de programación. Sigue estos pasos.

---

## 1. Instalar (solo la primera vez)

### En Windows
1. Descarga el archivo **`PMP-Windows.zip`**.
2. Haz clic derecho sobre él → **Extraer todo…**
3. Entra a la carpeta y haz **doble clic en `PMP.exe`**.
4. La primera vez Windows puede mostrar una pantalla azul que dice
   *"Windows protegió tu PC"*. Es normal (la app no está firmada).
   Haz clic en **"Más información"** y luego en **"Ejecutar de todas formas"**.

### En Mac
1. Descarga el `.zip` de tu Mac:
   - Si es un Mac moderno (M1, M2, M3…): **`PMP-macOS-AppleSilicon.zip`**
   - Si es un Mac más antiguo (Intel): **`PMP-macOS-Intel.zip`**
   - ¿No sabes cuál? Menú Apple () → *Acerca de este Mac*. Si dice
     "Chip Apple" usa el de Apple Silicon; si dice "Procesador Intel", el de Intel.
2. Haz doble clic en el `.zip` para descomprimirlo.
3. Dentro verás dos archivos. Haz **doble clic en `Lanzar PMP.command`**.
4. La primera vez Mac dirá que *"no se puede abrir porque es de un desarrollador
   no identificado"*. Es normal. Para abrirlo:
   - Haz **clic derecho** sobre `Lanzar PMP.command` → **Abrir** →
     en el aviso, vuelve a pulsar **Abrir**.
   - A partir de ahí, las siguientes veces abre con doble clic normal.

> Mantén juntos los dos archivos (`pmp` y `Lanzar PMP.command`) en la misma
> carpeta. Si separas uno, no funcionará.

---

## 2. Usar la herramienta

Al abrir se ve un menú. Te mueves con las **flechas ↑ ↓** y eliges con **Enter**.

### Generar la semana siguiente
1. Elige **📋 Generar semana siguiente**.
2. La herramienta **detecta sola** la última semana que hay en el Control y propone
   generar la **siguiente**. Solo tienes que **confirmar con Enter** — ya no
   escribes fechas. (Si alguna vez necesitas otra semana, responde que no y podrás
   indicarla a mano.)
3. La herramienta busca sola los dos Excel. Si encuentra varios, te deja elegir
   con las flechas.
4. Te pregunta si **alguien está ausente** esa semana. Marca con la **barra
   espaciadora** a quien corresponda (o ninguno) y pulsa **Enter**.
5. Aparece una **vista previa** de cómo queda la rotación. Revísala.
6. Confirma con **Enter**. Se actualiza el Control (copiando la semana tal cual y
   rotando solo a los encargados de Célula 3) y la Matriz de disponibilidad.
7. Al final te pregunta si quieres **además** un **cuadro resumen** aparte para
   compartir al equipo. Es opcional: pulsa Enter para omitirlo.

### Consultar disponibilidad nocturna
Elige **🌙 Consultar disponibilidad nocturna**, indica la semana, y te muestra
quién está de N2 (incidentes) y N3 (escalamiento).

---

## 3. ¿Dónde quedan los archivos?

La herramienta trabaja sobre **un solo archivo**: actualiza tu **Control** y tu
**Matriz** directamente, sobre el mismo Excel (no crea una copia nueva cada
semana). Cada semana se añade la siguiente sobre ese mismo archivo.

- **Control**: se actualiza el mismo `Control_Gestion_PMP....xlsx` que abriste.
- **Matriz**: se actualiza la misma `Matriz....xlsx`.
- **Cuadro resumen** (solo si lo pides, para compartir): `PMP_Semana.xlsx` —
  nombre fijo, se reescribe cada vez (no se acumulan copias).

> El guardado es **seguro**: si algo fallara a mitad, tu archivo **no queda
> dañado** (se conserva como estaba). Aun así, como ahora se modifica el archivo
> directamente, conviene que tu copia de respaldo esté en SharePoint/OneDrive.

Al terminar, la herramienta te ofrece abrir el Control para que lo revises antes
de subirlo a SharePoint.

## 4. Si algo no sale bien

- **No encuentra los Excel**: entra a **⚙ Configurar rutas de archivos** y
  selecciona los archivos a mano. Quedan recordados para la próxima vez.
- **Dice que "ya existe un bloque" o que la semana no es la siguiente**: revisa
  que la fecha sea la del lunes correcto, justo la semana posterior a la última
  que ya está en el Control.
- **Cualquier otra cosa**: cierra la ventana y vuelve a abrir; si persiste,
  avísale a quien te pasó la herramienta con una foto de lo que aparece.
