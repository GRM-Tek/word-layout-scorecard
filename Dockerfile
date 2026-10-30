# Builds the pinned engine's command-line tool from the submodule tree. Nothing is downloaded at
# run time; the only network use is cargo fetching crates while building.
#
#   scripts/build.sh            (checks the pin, then: docker build --build-arg WORDCRAFT_COMMIT=... .)
#
# Run it sealed: scripts/run-sealed.sh (no network, inputs read-only, no fonts unless mounted).

FROM rust:1-bookworm AS builder
ARG WORDCRAFT_COMMIT=unknown
WORKDIR /src/wordcraft
COPY wordcraft/ /src/wordcraft/
RUN cargo build --release --locked -p wordcraft-cli \
 && strip target/release/wordcraft-cli \
 && mkdir -p /licences \
 && cp LICENSE-MIT LICENSE-APACHE NOTICE ATTRIBUTION.md /licences/ \
 && cp assets/fonts/OFL-*.txt /licences/

FROM debian:bookworm-slim
ARG WORDCRAFT_COMMIT=unknown
LABEL org.opencontainers.image.title="studio-editor engine cli" \
      org.opencontainers.image.revision="${WORDCRAFT_COMMIT}" \
      org.opencontainers.image.licenses="MIT OR Apache-2.0"
COPY --from=builder /src/wordcraft/target/release/wordcraft-cli /usr/local/bin/wordcraft-cli
COPY --from=builder /licences/ /usr/share/doc/wordcraft/
RUN useradd --create-home --uid 1000 runner \
 && mkdir -p /in /out /usr/share/fonts \
 && chown runner:runner /out \
 && echo "${WORDCRAFT_COMMIT}" > /usr/share/doc/wordcraft/COMMIT
USER runner
WORKDIR /home/runner
ENTRYPOINT ["wordcraft-cli"]
CMD ["--help"]

# A second image, built only on request (scripts/visual-diff.sh), holding the pinned tree's own page-comparison
# tool (`cargo xtask visual-diff`) and the poppler tools it calls. Used once per checkpoint to cross-check the
# scorecard's pixel measure; never part of the engine image.
#   docker build --target xtask -t studio-editor/wordcraft-xtask:latest .
FROM builder AS xtask-builder
RUN cargo build --release --locked -p xtask

FROM debian:bookworm-slim AS xtask
RUN apt-get update && apt-get install -y --no-install-recommends poppler-utils && rm -rf /var/lib/apt/lists/*
COPY --from=xtask-builder /src/wordcraft/target/release/xtask /usr/local/bin/xtask
RUN useradd --create-home --uid 1000 runner && mkdir -p /in /out && chown runner:runner /out
USER runner
WORKDIR /home/runner
ENTRYPOINT ["xtask"]
