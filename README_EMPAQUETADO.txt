# Empaquetado Archivo Documental

## Archivos incluidos

- `ArchivoDocumental.spec`
- `compilar_exe.bat`

## Cómo usar

Copia estos 2 archivos en la raíz de tu proyecto, donde está:

```text
app.py
app/
activos/
```

La estructura debe verse así:

```text
archivo_documental_app/
├── app.py
├── ArchivoDocumental.spec
├── compilar_exe.bat
├── app/
│   ├── main.py
│   ├── services.py
│   ├── database.py
│   └── ...
└── activos/
    └── escudo.png
```

Luego ejecuta:

```text
compilar_exe.bat
```

Al terminar, el programa queda en:

```text
dist/ArchivoDocumental/ArchivoDocumental.exe
```

## LibreOffice Portable

Para que PDF funcione sin instalar LibreOffice en cada equipo, descarga LibreOffice Portable y copia su contenido aquí:

```text
dist/ArchivoDocumental/LibreOfficePortable/
```

Debe existir:

```text
dist/ArchivoDocumental/LibreOfficePortable/program/soffice.exe
```

## app.db

El archivo `app.db` debe crearse automáticamente junto al `.exe` si tu `database.py` usa ruta dinámica basada en `sys.executable`.

