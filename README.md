# Gestión de Órdenes — Calendario

Aplicación web para visualizar órdenes desde un archivo Excel en un calendario.
Días con órdenes al día → verde pastel. Días con órdenes atrasadas → rojo.

---

## Archivos del proyecto

```
gestion-ordenes/
├── index.html   ← Estructura HTML principal
├── style.css    ← Estilos (paleta verde/rojo pastel)
├── app.js       ← Lógica del calendario, Excel y alertas
└── README.md    ← Este archivo
```

---

## Cómo abrir en VS Code

1. Abre VS Code.
2. Menú **File → Open Folder** y selecciona la carpeta `gestion-ordenes`.
3. Instala la extensión **Live Server** (busca "Live Server" de Ritwick Dey en Extensions).
4. Clic derecho en `index.html` → **Open with Live Server**.
5. Se abrirá el navegador en `http://127.0.0.1:5500`.

> Sin Live Server también funciona: doble clic en `index.html` para abrirlo directo en el navegador.

---

## Formato del Excel

Tu archivo `.xlsx` debe tener una hoja con al menos estas 3 columnas
(el nombre exacto no importa, el sistema las detecta automáticamente):

| numero de orden | estado de orden | orden       |
|-----------------|-----------------|-------------|
| ORD-001         | Pendiente       | 15/05/2026  |
| ORD-002         | Completado      | 10/05/2026  |
| ORD-003         | En proceso      | 20/05/2026  |

**Formatos de fecha aceptados:**
- `DD/MM/YYYY` → ej: `20/05/2026`
- `YYYY-MM-DD` → ej: `2026-05-20`
- `DD-MM-YYYY` → ej: `20-05-2026`
- Fecha nativa de Excel (número serial)

**Estados que se consideran "completados"** (no generan alerta):
- Completado, Completada
- Entregado, Entregada
- Finalizado, Finalizada
- Cerrado, Cerrada
- Terminado, Terminada
- Done, Complete

Cualquier otro estado con fecha pasada se marca como **Atrasada**.

---

## Funcionalidades

- **Calendario mensual** con navegación mes anterior/siguiente.
- **Color verde** en días con órdenes al día.
- **Color rojo** en días con órdenes atrasadas.
- **Clic en un día** → detalle de las órdenes de ese día en el panel derecho.
- **Panel de alertas** → lista de días atrasados con sus órdenes.
- **Tabla completa** de todas las órdenes con estado y alerta.
- **Drag & Drop** para soltar el Excel directamente en el área de carga.
- **Datos de ejemplo** para probar sin un archivo real.
- Diseño **responsive** (en móvil las columnas se apilan).

---

## Dependencias (CDN, no requieren instalación)

- [SheetJS / xlsx](https://sheetjs.com/) — lectura de archivos Excel
- [Tabler Icons](https://tabler.io/icons) — iconografía
