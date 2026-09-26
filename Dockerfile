# ---------- build the .accdb Java engine ----------
FROM maven:3.9-eclipse-temurin-17 AS jbuild
WORKDIR /jb
COPY accdb_engine/pom.xml .
COPY accdb_engine/src ./src
RUN mvn -q -DskipTests package

# ---------- python runtime ----------
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends openjdk-17-jre-headless \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN python -m pip install --upgrade pip setuptools wheel \
    && pip install -r requirements.txt

COPY . .
COPY --from=jbuild /jb/target/accdbtool.jar /app/accdbtool.jar
RUN chmod +x /app/start.sh

EXPOSE 8080

CMD ["sh", "/app/start.sh"]
