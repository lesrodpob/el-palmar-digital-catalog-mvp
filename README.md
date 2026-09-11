# El Palmar — Pro Catalog (Excel Data Source)

React/Vite prototype of the El Palmar digital catalog using the uploaded Excel report as its current data source.

## Data flow

`public/data/productos.xlsx` → `src/main.jsx` (`xlsx` parser) → React UI

The data layer is intentionally isolated so the Excel source can later be replaced by the informatician's API endpoint without redesigning the frontend.

## Excel fields used

- ID
- Código Barra
- Categoría
- Familia
- Nombre Producto
- Precio
- MATRIZ (used as current stock)

## Run locally

```bash
npm install
npm run dev
```

Then open the local Vite URL.

## Important

This prototype reads the Excel file in the browser. If the repository is public, the Excel file is also publicly accessible. Before production, replace the Excel source with the planned API endpoint or a safe public-data feed.
