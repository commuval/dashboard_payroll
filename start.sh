#!/bin/bash

# Streamlit App auf DigitalOcean starten
streamlit run app.py --server.port=$PORT --server.address=0.0.0.0 --server.headless=true 