# FinancieroData · Dashboard Financiero y Contable Ejecutivo

Dashboard analítico interactivo de alto rendimiento para la visualización, control y toma de decisiones sobre balances de comprobación y saldos de cuentas mensualizados (`.xlsx`, `.xls`, `.csv`).

Desarrollado para el repositorio: **[Efraingon/maxmaxdata-dashboard](https://github.com/Efraingon/maxmaxdata-dashboard)**

---

## 🎯 Objetivo de la Aplicación

El objetivo principal de **FinancieroData Dashboard** es brindar a directores financieros, contadores y auditores una herramienta visual, ágil y 100% privada para:
1. **Analizar la evolución mensual** de los diferentes grupos contables (Activos, Pasivos, Patrimonio, Ingresos, Costos, Gastos).
2. **Evaluar la concentración y distribución patrimonial** mediante gráficos circulares y jerárquicos (Treemaps).
3. **Identificar cuentas clave y anomalías** a través de ordenamiento por saldos absolutos y conclusiones automáticas.
4. **Facilitar la carga de datos** proporcionando una plantilla oficial estandarizada en Excel lista para ser completada.
5. **Garantizar la privacidad absoluta**: Todo el procesamiento se realiza en el cliente (navegador del usuario) sin enviar datos contables a servidores externos.

---

## 🛠️ Herramientas y Tecnologías Utilizadas

| Categoría | Tecnología / Librería | Versión | Propósito |
| :--- | :--- | :--- | :--- |
| **Core UI** | [React](https://react.dev/) | `^18.3.1` | Biblioteca base para componentes reactivos y estado de la aplicación. |
| **Render Engine** | [React DOM](https://react.dev/) | `^18.3.1` | Renderizado eficiente en el DOM. |
| **Build & Dev Tool** | [Vite](https://vitejs.dev/) | `^5.4.11` | Empaquetador ultrarrápido con Hot Module Replacement (HMR). |
| **Vite Plugin** | `@vitejs/plugin-react` | `^4.3.4` | Integración y optimización de JSX con Fast Refresh. |
| **Visualización / Gráficos** | [Recharts](https://recharts.org/) | `^2.13.3` | Gráficos interactivos SVG (Barras apiladas, Donut/Pie, Treemap, Barras horizontales). |
| **Procesamiento de Archivos** | [SheetJS (XLSX)](https://sheetjs.com/) | `^0.18.5` | Lectura, parseo y exportación en memoria de hojas `.xlsx`, `.xls` y `.csv`. |
| **Iconografía** | [Lucide React](https://lucide.dev/) | `^0.468.0` | Conjunto de iconos vectoriales limpios y coherentes. |
| **Estilos** | CSS3 Moderno / Vanilla CSS | — | Diseño responsivo, paleta ejecutiva, variables HSL, scroll optimizado y layouts CSS Grid/Flexbox. |

---

## 📂 Esquema de Archivos del Proyecto

```text
maxmaxdata-dashboard/
├── .gitignore                         # Archivos y carpetas ignorados por Git (node_modules, dist, etc.)
├── index.html                         # Plantilla HTML inicial
├── package.json                       # Configuración de dependencias y scripts de Vite
├── package-lock.json                  # Árbol de dependencias bloqueadas
├── README.md                          # Documentación completa del proyecto
├── dist/                              # Compilación de producción (generada por `npm run build`)
│   ├── assets/                        # CSS y JS minificados
│   └── index.html                     # Archivo HTML optimizado
└── src/                               # Código fuente principal
    ├── main.jsx                       # Componente raíz, lógica de estado, parseo Excel y visualizaciones
    └── styles.css                     # Sistema de diseño, temas de color, KPIs, sidebar y tablas
```

---

## ⚙️ Arquitectura y Componentes Clave (`src/main.jsx`)

1. **Motor de Ingesta (`parseWorkbook`)**:
   - Detección dinámica de la fila de encabezado que contiene la columna `Cuenta`.
   - Extracción de metadatos institucionales (`Empresa`, `Período`, `RIF`).
   - Normalización de números con separador decimal europeo o latinoamericano.
2. **Generador de Plantilla (`downloadTemplate`)**:
   - Genera una hoja Excel `.xlsx` estructurada con cuentas del plan contable estándar agrupadas del 1 al 6.
3. **Agrupación Contable Dinámica (`getAccountGroupCode` y `getGroupName`)**:
   - Agrupa cuentas por clase/grupo contable mayor (`1 - Activos`, `2 - Pasivos`, `3 - Patrimonio`, `4 - Ingresos`, `5 - Costos`, `6 - Gastos`, etc.).
   - Mantiene orden secuencial numérico (`1, 2, 3...`).
4. **Vistas / Pestañas**:
   - **Resumen General**: KPIs principales, Evolución Mensual por Grupo, Distribución por Grupo, Top Cuentas Absolutas, Mapa de Cuentas (Treemap) y Conclusiones automáticas.
   - **Análisis de Categorías**: Jerarquía de cuentas y comparativa de grupos principales.
   - **Detalle de Registros**: Tabla interactiva agrupada con totales, ordenamiento por columnas, paginación y exportación.

---

## 🚀 Instalación y Puesta en Marcha

### Prerrequisitos
- [Node.js](https://nodejs.org/) (versión 18 o superior recomendada).
- Gestor de paquetes `npm`.

### 1. Clonar o descargar el repositorio
```bash
git clone https://github.com/Efraingon/maxmaxdata-dashboard.git
cd maxmaxdata-dashboard
```

### 2. Instalar dependencias
```bash
npm install
```

### 3. Ejecutar en modo desarrollo
```bash
npm run dev
```
Abre en tu navegador: `http://localhost:5173/`

### 4. Compilar para producción
```bash
npm run build
npm run preview
```

---

## 📤 Pasos para Subir al Repositorio GitHub (`Efraingon/maxmaxdata-dashboard`)

Si estás subiendo el proyecto por primera vez:

```bash
# 1. Inicializar git en la carpeta
git init

# 2. Agregar todos los archivos
git add .

# 3. Crear el primer commit
git commit -m "feat: initial commit of FinancieroData dashboard"

# 4. Establecer la rama principal
git branch -M main

# 5. Conectar con el repositorio remoto de GitHub
git remote add origin https://github.com/Efraingon/maxmaxdata-dashboard.git

# 6. Subir los cambios
git push -u origin main
```

---

## 📄 Licencia
Este proyecto es de uso privado / corporativo para análisis financiero y contable.
