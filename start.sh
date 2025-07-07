#!/bin/bash

# Port aus Umgebungsvariable oder Standard verwenden
PORT=${PORT:-8080}

# Streamlit App mit expliziten Parametern starten
exec streamlit run app.py \
  --server.port=$PORT \
  --server.address=0.0.0.0 \
  --server.headless=true \
  --server.enableCORS=false \
  --server.enableXsrfProtection=false \
  --server.maxUploadSize=200 