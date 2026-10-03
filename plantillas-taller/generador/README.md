# Generador de las plantillas de taller

Tres plantillas de PowerPoint construidas con la piel Super Humanos del design system Campus Pivot.

- `e1.js` → Encuentro 1 · concepto y práctica (90 minutos)
- `consultorio.js` → Consultorio (60 minutos)
- `e2.js` → Encuentro 2 · aplicación (90 minutos)
- `brand.js` → colores, tipografía, íconos del sistema y piezas (cinta, barra de chevrones, tarjeta, bloque de prompt)
- `common.js` → portada común y construcción del archivo

Para regenerar:

```
npm install
NODE_PATH=$PWD/node_modules node e1.js
NODE_PATH=$PWD/node_modules node consultorio.js
NODE_PATH=$PWD/node_modules node e2.js
```

Los archivos salen en `out/`. Fuentes requeridas en el equipo que presenta: Barlow Condensed (700) y Atkinson Hyperlegible Next (400 y 700), ambas en Google Fonts.
