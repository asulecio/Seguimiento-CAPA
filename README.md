# Seguimiento Post-CAPA · Blue Medical

Dashboard de consulta del seguimiento de nuevos ingresos después de la capacitación CAPA.

- La página es **solo de consulta** y pide la **contraseña del equipo**.
- Los datos están en `datos.enc.json`, **cifrados** (AES-256-GCM con clave derivada de la contraseña). Sin la contraseña no se pueden leer.
- Los datos se generan a partir del Excel de formadores con `actualizar_datos.py`:

```
python3 actualizar_datos.py Seguimiento_Post-CAPA_formadores.xlsx datos.enc.json
```

El script pide la contraseña y crea `datos.enc.json`; después se sube ese archivo a este repositorio para actualizar la página.

No subas el Excel ni ningún archivo con datos sin cifrar a este repositorio: es público.
