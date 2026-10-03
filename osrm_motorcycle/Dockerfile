Rename-Item .\osrm_motorcycle\Dockerfile.txt DockerfileFROM debian:bookworm-slim AS downloader

RUN apt-get update \
    && apt-get install -y --no-install-recommends curl ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /data

RUN curl -L \
    https://download.geofabrik.de/asia/taiwan-latest.osm.pbf \
    -o /data/taiwan.osm.pbf


FROM ghcr.io/project-osrm/osrm-backend:latest AS builder

WORKDIR /data

COPY --from=downloader /data/taiwan.osm.pbf /data/taiwan.osm.pbf
COPY data/motorcycle.lua /data/motorcycle.lua

RUN osrm-extract -p /data/motorcycle.lua /data/taiwan.osm.pbf \
    && osrm-partition /data/taiwan.osrm \
    && osrm-customize /data/taiwan.osrm


FROM ghcr.io/project-osrm/osrm-backend:latest

WORKDIR /data

COPY --from=builder /data /data

CMD ["sh", "-c", "osrm-routed --algorithm mld --ip 0.0.0.0 --port ${PORT:-10000} --mmap /data/taiwan.osrm"]