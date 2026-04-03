# Linux-only image for headless slide export (LibreOffice + Poppler).
# On Windows, run the project natively: the CLI uses PowerPoint COM automatically
# and does not require LibreOffice.

FROM python:3.12-slim-bookworm

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-nogui \
        poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY pyproject.toml README.md ./
COPY src ./src

RUN pip install --no-cache-dir -U pip \
    && pip install --no-cache-dir -e .

WORKDIR /work
ENTRYPOINT ["python", "-m", "note2video"]
CMD ["--help"]
