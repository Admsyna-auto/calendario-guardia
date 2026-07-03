# Guardia Administrativa — Riiing / Syna S.A.
## Proyecto de scheduling y visualización de turnos

---

## Estructura del proyecto

```
guardia/
├── index.html              ← Widget web (abrir directo en browser o deploy en GitHub Pages)
├── README.md               ← Este archivo
└── scripts/
    ├── build_cuadro.py     ← Excel: una columna por persona dentro de cada día (formato imagen)
    ├── build_franja.py     ← Excel: franjas horarias por día (versión anterior)
    ├── build_schedule.py   ← Excel: calendario mensual (vista tabla)
    └── build_semanal.py    ← Excel: vista semanal (hoja por semana)
```

---

## Horario vigente (Julio 2026)

| Día     | Ana     | Franco  | Yazmin  | Martina | Victoria |
|---------|---------|---------|---------|---------|----------|
| Lunes   | FRANCO  | 13-22   | 12-21   | 9-17    | 14-22    |
| Martes  | 13-22   | FRANCO  | 12-21   | 9-17    | 14-22    |
| Miércoles | 12-21 | FRANCO  | FRANCO  | 9-17    | 14-22    |
| Jueves  | 12-22   | 12-22   | FRANCO  | 9-17    | FRANCO   |
| Viernes | 12-22   | 9-19    | 9-19    | FRANCO  | FRANCO   |
| Sábado  | 12-22   | 9-19    | 9-19    | FRANCO  | FRANCO   |
| Domingo | FRANCO  | 13-22   | 12-22   | FRANCO  | 10-18    |

### Francos semanales
- **Ana**: Domingo + Lunes
- **Franco**: Martes + Miércoles
- **Yazmin**: Miércoles + Jueves
- **Martina**: Viernes + Sábado + Domingo
- **Victoria**: Jueves + Viernes + Sábado

### Contratos
- Ana, Franco, Yazmin: 48 hs/semana — 1 h descanso/día
- Martina, Victoria: 32 hs/semana — 30 min descanso/día

---

## Widget web (`index.html`)

Archivo HTML standalone sin dependencias externas. Funciona offline.

### Características
- Grilla semanal: franjas horarias en filas, días en columnas, **una columna por persona**
- Navegación ‹ › entre semanas
- Hora actual marcada en amarillo (se refresca cada 30 s)
- Columna del día actual resaltada
- Pulso animado en personas activas ahora mismo
- Módulo de configuración (⚙): editar turnos con guardado en `localStorage`
- Feriado 9/7 marcado automáticamente
- Deploy: subir `index.html` como `index.html` a GitHub Pages

### Deploy en GitHub Pages
```bash
# Crear repo en github.com/admsyna-auto/guardia
# Subir index.html → Settings → Pages → main / root → Save
# URL: https://admsyna-auto.github.io/guardia
```

---

## Scripts Python (generadores de Excel)

Requieren: `pip install openpyxl`

```bash
python scripts/build_cuadro.py    # genera guardia_julio_2026_cuadro.xlsx
python scripts/build_schedule.py  # genera guardia_julio_2026.xlsx
```

### Datos del schedule en los scripts

```python
SCHEDULE = {
    # dow: 0=Lun, 1=Mar, 2=Mié, 3=Jue, 4=Vie, 5=Sáb, 6=Dom
    "Ana":      {0:None,   1:"13-22",2:"12-21",3:"12-22",4:"12-22",5:"12-22",6:None},
    "Franco":   {0:"13-22",1:None,   2:None,   3:"12-22",4:"9-19", 5:"9-19", 6:"13-22"},
    "Yazmin":   {0:"12-21",1:"12-21",2:None,   3:None,   4:"9-19", 5:"9-19", 6:"12-22"},
    "Martina":  {0:"9-17", 1:"9-17", 2:"9-17", 3:"9-17", 4:None,   5:None,   6:None},
    "Victoria": {0:"14-22",1:"14-22",2:"14-22",3:None,   4:None,   5:None,   6:"10-18"},
}
```

### Colores por persona
| Persona  | Fondo (hex) | Texto (hex) |
|----------|-------------|-------------|
| Ana      | B5D4F4      | 0C447C      |
| Franco   | F5C4B3      | 712B13      |
| Yazmin   | C0DD97      | 27500A      |
| Martina  | FAC775      | 633806      |
| Victoria | CECBF6      | 3C3489      |

---

## Contexto del negocio

- **Empresa**: Riiing (operada por Syna S.A.)
- **Área**: Administración — Guardia Administrativa
- **Equipo**: 5 personas (Ana, Franco, Yazmin, Martina, Victoria)
- **Horario guardia**: Lun–Sáb 9:00–22:00 / Dom 10:00–22:00
- **Días pico**: Viernes y Sábados (cobertura mínima 3 personas 12–22)
- **Feriado marcado**: 9 de julio — Día de la Independencia

---

## Ideas para extender el proyecto

- [ ] Soporte multi-mes (no solo julio 2026)
- [ ] Vista de un solo día con detalle de cobertura por franja
- [ ] Notificaciones (quién empieza/termina en la próxima hora)
- [ ] Panel admin con login para editar horarios de forma persistente (backend)
- [ ] Exportar a PDF / imprimir desde el widget
- [ ] Integración con Google Calendar o WhatsApp Business
