# Memoria del Proyecto - Archivo Web

## Estado Actual
La aplicación ha sido migrada de una versión de escritorio a una plataforma web moderna con arquitectura Flask + SPA (Vanilla JS).

## Mejoras Recientes (Mayo 2026)
- **Dashboard Principal:** Panel de control con estadísticas en tiempo real y gráficos de tendencia (Chart.js).
- **Modo Oscuro:** Implementación de temas dinámicos (Claro/Oscuro).
- **Paginación:** Los registros ahora se cargan de forma paginada desde el servidor para optimizar el rendimiento.
- **Anexos y Adjuntos:** Soporte para subir, descargar y eliminar archivos (PDF, Imágenes) asociados a cada registro documental.
- **Registro de Auditoría:** Interfaz para que los administradores supervisen todas las acciones realizadas en el sistema.
- **Sincronización GitHub:** Proyecto inicializado en Git y subido exitosamente a [https://github.com/aranjara/Programa-Gestion-Documental](https://github.com/aranjara/Programa-Gestion-Documental).

## Arquitectura Técnica
- **Control de Versiones:** Git (GitHub).
- **Backend:** Flask (Python) con SQLite.
- **Frontend:** HTML5, CSS3 (Glassmorphism), JavaScript Vanilla.
- **Seguridad:** Hash de contraseñas con PBKDF2-SHA256, sesiones protegidas y roles de usuario (admin/normal).
- **Reportes:** Generación de FUID en Excel/Word/PDF y Rótulos de caja/carpeta.

## Credenciales por Defecto (Entorno de Desarrollo)
- **Admin:** admin / 123456 (o la configurada en sesión).
- **Base de datos:** `app.db`.
- **Archivos:** Almacenados en la carpeta `uploads/`.

## Project Overview
This project is a web-based archive management system ("Programa Archivo Web"). It was originally a desktop application built with Tkinter and has been migrated to a modern web architecture using Flask.

## Tech Stack
- **Backend**: Python with Flask
- **Database**: SQLite (`app.db`)
- **Frontend**: HTML5, CSS3 (Vanilla), JavaScript, GSAP (AI Skills installed in `skill.md/`)
- **Deployment/Packaging**: PyInstaller (indicated by `.spec` and `compilar_exe.bat`)

## Key Features
- Document and archive management.
- User authentication and session management.
- Generation of reports (PDF, Excel, Word - mentioned in migration history).
- Responsive web interface.

## Historical Context (Key Milestones)
- **Migration**: Successfully converted the core business logic from Tkinter to Flask API endpoints.
- **Authentication**: Implemented and refined login error handling and password recovery mechanisms.
- **Development**: Active development is focusing on enhancing the web UI and ensuring robust backend verification.

## Architecture Notes
- `app_web.py`: Main entry point for the Flask application.
- `static/`: Contains frontend assets (`js/app.js`, `css/style.css`, `index.html`).
- `templates/`: Flask HTML templates (if used, though `index.html` is in `static`).
- `app.db`: SQLite database for persistent storage.

## User Preferences & Decisions
- (To be updated as the conversation progresses)
