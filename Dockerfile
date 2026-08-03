# mamaXO Document Factory — container for any Docker host (Hugging Face Spaces, Render, Railway, Fly)
FROM python:3.12-slim

# system libraries WeasyPrint needs to render PDFs
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf-2.0-0 libcairo2 libffi8 \
    shared-mime-info fonts-dejavu-core \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PORT=7860
EXPOSE 7860
# gunicorn for a small team; 2 workers is plenty
CMD ["gunicorn", "-b", "0.0.0.0:7860", "-w", "2", "--timeout", "120", "app:app"]
