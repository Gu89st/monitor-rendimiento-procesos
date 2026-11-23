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
   APP_PASSWORD=una_contrasena_aleatoria_larga
   ```

   Sustituye `APP_PASSWORD` por una contraseña aleatoria de al menos 16 caracteres. Para habilitar el análisis con IA, agrega también `GEMINI_API_KEY`. `GOOGLE_API_KEY` es una alternativa. `GEMINI_MODEL` y `GEMINI_FALLBACK_MODELS` son opcionales.

3. Inicia la aplicación:

   ```powershell
   python -m streamlit run app.py
   ```

## Desplegar en Streamlit Community Cloud

> **Importante:** la URL de Streamlit Community Cloud es pública, aunque el repositorio sea privado. La app exige `APP_PASSWORD` para acceder y permite cargar y eliminar datos de Supabase; configura una contraseña aleatoria de al menos 16 caracteres y compártela solo con usuarios autorizados.

1. Sube este proyecto a un repositorio de GitHub.
2. En Streamlit Community Cloud, crea una app desde ese repositorio y selecciona `app.py` como archivo principal.
3. En **Settings → Secrets**, configura `APP_PASSWORD`, `SUPABASE_URL` y, si utilizarás el análisis con IA, `GEMINI_API_KEY` (o `GOOGLE_API_KEY`). No subas `.env` ni claves al repositorio.
