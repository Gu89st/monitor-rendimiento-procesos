# Monitor de Rendimiento de Procesos

Aplicación Streamlit para analizar ejecuciones de procesos, visualizar tiempos de demora y consultar datos almacenados en Supabase.

## Ejecutar localmente

1. Instala las dependencias:

   ```powershell
   python -m pip install -r requirements.txt
   ```

2. Crea un archivo `.env` junto a `app.py` y configura al menos:

   ```dotenv
   SUPABASE_URL=tu_url_de_conexion_a_supabase
   ```

   Para habilitar el análisis con IA, agrega también `GEMINI_API_KEY`. `GOOGLE_API_KEY` es una alternativa. `GEMINI_MODEL` y `GEMINI_FALLBACK_MODELS` son opcionales.

3. Inicia la aplicación:

   ```powershell
   python -m streamlit run app.py
   ```

## Desplegar en Streamlit Community Cloud

> **Importante:** la app todavía no requiere autenticación y permite cargar y eliminar datos de Supabase. Streamlit Community Cloud publica la app; que el repositorio de GitHub sea privado no hace privada la aplicación. No la despliegues con datos reales hasta agregar control de acceso o elegir un alojamiento privado.

1. Sube este proyecto a un repositorio de GitHub.
2. En Streamlit Community Cloud, crea una app desde ese repositorio y selecciona `app.py` como archivo principal.
3. En **Settings → Secrets**, configura `SUPABASE_URL` y, si utilizarás el análisis con IA, `GEMINI_API_KEY` (o `GOOGLE_API_KEY`). No subas `.env` ni claves al repositorio.
