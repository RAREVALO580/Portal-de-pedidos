# 🚚 Portal de Traslados | Seguimiento a Tienda (Macropay Retail)

> **Dashboard Interactivo de Business Intelligence & Monitoreo Logístico CEDI → Sucursales basado en SAP ERP MM (ME2W y ME80FN)**

---

## 🌐 Demo Pública en Vivo (GitHub Pages)
Una vez activado **GitHub Pages** en el repositorio, la aplicación estará disponible públicamente en:
👉 `https://<tu-usuario-github>.github.io/<nombre-repositorio>/`

---

## 📊 Descripción del Proyecto
El **Portal de Traslados** es una plataforma web interactiva de monitoreo, rastreo y recepción física de mercancía para la red de 635 tiendas Macropay. Inspirada en la experiencia de usuario (UX) de plataformas de entrega de última milla (estilo **Uber Eats / Amazon Logistics**), permite a los supervisores de sucursal y coordinadores de tráfico:

* Identificar pedidos de traslado en camino en tiempo real.
* Rastrear el avance a través de un **Stepper dinámico de 4 hitos** sincronizado con apuntes contables reales de SAP.
* Operar un **Checklist físico de recepción** durante la descarga del transporte con formato condicional y opción de impresión formal.

---

## ⚡ Origen de Datos y Reglas de Negocio SAP MM
El modelo consolida dos transacciones estándar de SAP MM:
1. **SAP ME2W:** Cabecera y detalle de pedidos de traslado activos (`EBELN`, `EBELP`, `MATNR`, `TXZ01`, `WERKS`, `RESWK`, `MENGE`, `AEDAT`, `LOEKZ IS NULL`).
2. **SAP ME80FN:** Vista de historial de pedidos (`EKBE`) con clases de movimiento contable (`351`, `641` para despacho en CEDI; `101` para recepción física en tienda).

### Fórmulas de Transformación
* **Cant_Pedida** = `MENGE` de ME2W.
* **Cant_Despachada** = `SUM(CASE WHEN BWART IN ('351', '641') THEN MENGE WHEN BWART IN ('352', '642') THEN (MENGE * -1) ELSE 0 END)`
* **Cant_Recibida** = `SUM(CASE WHEN BWART = '101' THEN MENGE WHEN BWART IN ('102', '104') THEN (MENGE * -1) ELSE 0 END)`
* **Cant_En_Transito** = `GREATEST((Cant_Despachada - Cant_Recibida), 0)`

---

## 🚀 Estructura de Archivos
* **`index.html` / `portal_traslados.html`:** Aplicación interactiva Single Page Application (SPA).
* **`portal_data.js`:** Payload de datos consolidado para carga instantánea offline/online.
* **`dashboard_data.json`:** Dataset estructurado en JSON.
* **`actualizar_datos.bat` / `actualizar_datos.ps1`:** Script de extracción y actualización automática en memoria desde Excel SAP.
* **`abrir_dashboard.bat`:** Acceso directo local para abrir en Windows.

---

## 🛠️ Tecnologías Utilizadas
* HTML5 / CSS3 / JavaScript Moderno (Vanilla ES6+ sin dependencias externas pesadas)
* Google Fonts (*Plus Jakarta Sans*, *JetBrains Mono*)
* PowerShell / Excel COM Automation para ETL de alta velocidad
* GitHub Pages para despliegue y visualización pública
