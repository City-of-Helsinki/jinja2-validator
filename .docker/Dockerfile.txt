# ==============================
# Base image
# ==============================
FROM registry.access.redhat.com/ubi9/python-312 AS appbase

USER root
WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Install system dependencies
RUN dnf update -y \
    && dnf install -y gettext \
    && dnf clean all

# Install Python dependencies
COPY requirements.txt .
RUN pip install -U pip setuptools wheel \
    && pip install --no-cache-dir -r requirements.txt

# ==============================
# Final image
# ==============================
FROM appbase AS production

WORKDIR /app

# Copy application code
COPY . .

# Ensure entrypoint is executable
RUN chmod +x .docker/docker-entrypoint.sh

# Expose the Flask/Gunicorn port (matches docker-entrypoint.sh default)
EXPOSE 5001/tcp

ENTRYPOINT ["/app/.docker/docker-entrypoint.sh"]
