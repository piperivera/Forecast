# ForeCast

## Estructura recomendada del proyecto

Sí, para este tipo de dashboard es mejor separar el proyecto por responsabilidad en lugar de mantener toda la lógica en `index.html`.

```text
ForeCast/
  index.html
  src/
    styles/
      main.css
    scripts/
      state.js          # variables globales: ALL_DATA, CURRENT_USER, TRM
      config.js         # ROLES, COLORS, KEY_MAP, MES_LABELS
      auth.js           # MSAL, login, showUserBadge, switchView
      data-loader.js    # loadFolderFromSharePoint, parseXlsx, fetchTRM
      render/
        kpis.js         # renderGerencia, renderKPIs
        charts.js       # renderEvoChart, renderDonut, evoSVG
        tables.js       # buildTable, renderGerenciaEstadoTables
        director.js     # renderDirector
        ejecutivo.js    # renderEjecutivo
        divisas.js      # renderDivisas
        marcas.js       # renderMarcas
      utils.js          # abr, parseMonto, parseFecha, toCOP, fmtCOP
      main.js           # DOMContentLoaded, navegación, renderAll
```

## Orden sugerido de migración

1. Mover CSS embebido a `src/styles/main.css`.
2. Mover JS embebido a `src/scripts/main.js`.
3. Extraer constantes a `config.js` y estado a `state.js`.
4. Extraer carga/parseo a `data-loader.js`.
5. Dividir render por módulos en `src/scripts/render/`.
6. Dejar `main.js` como orquestador (eventos + render global).

## Reglas prácticas

- `index.html` debe tener solo estructura y referencias a CSS/JS.
- Los módulos de `render/` no deben modificar estado global directamente.
- `utils.js` solo contiene funciones puras reutilizables.
- `state.js` centraliza el estado mutable.

Esta organización mejora mantenimiento, pruebas y escalabilidad.
