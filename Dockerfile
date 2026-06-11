# syntax=docker/dockerfile:1
#
# Multi-stage pipeline (PLAN.md §13 walking skeleton, Docker-first):
#
#   go-test    – unit tests for the pure Go packages (race detector)
#   php-test   – shim test suite against the fake ABI
#   generate   – `frankenphp extension-init` turns bridge directives into CGO code
#   build      – xcaddy compiles FrankenPHP with the extension baked in
#   runner     – production FrankenPHP image with easy-excel + the PHP shim
#
#   docker build --target=go-test .        # run Go tests
#   docker build --target=php-test .       # run PHP shim tests
#   docker build -t frankenphp-easy-excel . # full custom FrankenPHP

ARG PHP_VERSION=8.4
ARG GO_VERSION=1.26

# --- tests --------------------------------------------------------------------

FROM golang:${GO_VERSION} AS go-test
WORKDIR /src
COPY extension/go.mod extension/go.sum ./
RUN go mod download
COPY extension/ .
RUN go vet ./registry/... ./limits/... ./exio/... ./compat/... ./core/... \
 && go test -race ./registry/... ./limits/... ./exio/... ./compat/... ./core/...

FROM php:${PHP_VERSION}-cli AS php-test
WORKDIR /app
COPY php/ php/
RUN php php/tests/run.php

# --- extension generation -------------------------------------------------------

FROM dunglas/frankenphp:builder-php${PHP_VERSION} AS generate
ARG PHP_VERSION
RUN apt-get update && apt-get install -y --no-install-recommends git \
 && rm -rf /var/lib/apt/lists/* \
 && git clone --depth=1 --branch=PHP-${PHP_VERSION} https://github.com/php/php-src /opt/php-src
WORKDIR /go/src/easy-excel/extension
COPY extension/ .
RUN GEN_STUB_SCRIPT=/opt/php-src/build/gen_stub.php frankenphp extension-init easy_excel.go \
 && ls -la build/

# --- full FrankenPHP build --------------------------------------------------------

FROM generate AS build
COPY --from=caddy:builder /usr/bin/xcaddy /usr/bin/xcaddy
ENV CGO_ENABLED=1 \
    XCADDY_SETCAP=1 \
    XCADDY_GO_BUILD_FLAGS="-ldflags='-w -s' -tags=nobadger,nomysql,nopgx"
RUN CGO_CFLAGS="$(php-config --includes)" \
    CGO_LDFLAGS="$(php-config --ldflags) $(php-config --libs)" \
    xcaddy build \
      --output /usr/local/bin/frankenphp-easy-excel \
      --with github.com/dunglas/frankenphp/caddy \
      --with github.com/ronisaha/easy-excel/extension/build=./build \
 && /usr/local/bin/frankenphp-easy-excel version

# --- runtime ------------------------------------------------------------------------

FROM dunglas/frankenphp:1-php${PHP_VERSION} AS runner
COPY --from=build /usr/local/bin/frankenphp-easy-excel /usr/local/bin/frankenphp
# the PhpSpreadsheet-compatible shim, installable via composer path-repository
COPY php/ /opt/easy-excel/php/
LABEL org.opencontainers.image.title="frankenphp-easy-excel" \
      org.opencontainers.image.description="FrankenPHP with the easy-excel (Go/excelize) spreadsheet extension" \
      org.opencontainers.image.source="https://github.com/ronisaha/easy-excel"
