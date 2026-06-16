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

ARG PHP_VERSION=8.5
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

# --- compat surface gate ------------------------------------------------------
# Reflects a real phpoffice/phpspreadsheet and diffs its public surface against
# the Compat layer. compat-surface-deps installs PhpSpreadsheet; compat-surface
# runs the diff against the committed baseline (php/.compat-surface.json) — a
# *new* uncovered class/method/constant fails the build instead of surfacing at
# runtime. Refresh the baseline deliberately with:
#   docker build --target=compat-surface-deps -t cs . && \
#   docker run --rm cs php php/tools/compat-surface-diff.php --update-baseline=- \
#     > php/.compat-surface.json
FROM php:${PHP_VERSION}-cli AS compat-surface-deps
WORKDIR /app
COPY --from=mlocati/php-extension-installer:latest /usr/bin/install-php-extensions /usr/local/bin/
RUN install-php-extensions gd zip intl xsl mbstring
COPY --from=composer:2 /usr/bin/composer /usr/local/bin/composer
ENV EASY_EXCEL_ALIAS=off COMPOSER_ALLOW_SUPERUSER=1
COPY php/ php/
RUN composer --working-dir=php require --no-interaction --no-progress --no-audit phpoffice/phpspreadsheet

FROM compat-surface-deps AS compat-surface
RUN php php/tools/compat-surface-diff.php --members --baseline=php/.compat-surface.json

# --- extension generation -------------------------------------------------------

FROM dunglas/frankenphp:builder-php${PHP_VERSION} AS generate
ARG PHP_VERSION
# libbrotli-dev: the caddy-cbrotli module (br encoder in the stock Caddyfile)
RUN apt-get update && apt-get install -y --no-install-recommends git libbrotli-dev \
 && rm -rf /var/lib/apt/lists/* \
 && git clone --depth=1 --branch=PHP-${PHP_VERSION} https://github.com/php/php-src /opt/php-src
WORKDIR /go/src/easy-excel/extension
COPY extension/ .
# generates easy_excel.c/.h, easy_excel_arginfo.h, easy_excel_generated.go and
# easy_excel.stub.php in-place next to the bridge source
RUN GEN_STUB_SCRIPT=/opt/php-src/build/gen_stub.php frankenphp extension-init easy_excel.go \
 && test -f easy_excel_generated.go && test -f easy_excel_arginfo.h

# --- full FrankenPHP build --------------------------------------------------------

FROM generate AS build
COPY --from=caddy:builder /usr/bin/xcaddy /usr/bin/xcaddy
ENV CGO_ENABLED=1 \
    XCADDY_SETCAP=1 \
    XCADDY_GO_BUILD_FLAGS="-ldflags='-w -s' -tags=nobadger,nomysql,nopgx"
# -D_GNU_SOURCE: zend_operators.h uses memrchr, hidden behind glibc's GNU extensions
# cbrotli: module parity with the official frankenphp binary —
# the runtime image's stock Caddyfile uses the br encoder, so without cbrotli
# the server refuses to start
RUN CGO_CFLAGS="$(php-config --includes) -D_GNU_SOURCE" \
    CGO_LDFLAGS="$(php-config --ldflags) $(php-config --libs)" \
    xcaddy build \
      --output /usr/local/bin/frankenphp-easy-excel \
      --with github.com/dunglas/frankenphp/caddy \
      --with github.com/dunglas/caddy-cbrotli \
      --with github.com/xiidea/easy-excel/extension=/go/src/easy-excel/extension \
 && /usr/local/bin/frankenphp-easy-excel version

# --- runtime ------------------------------------------------------------------------

FROM dunglas/frankenphp:1-php${PHP_VERSION} AS runner
COPY --from=build /usr/local/bin/frankenphp-easy-excel /usr/local/bin/frankenphp
# the PhpSpreadsheet-compatible shim, installable via composer path-repository
COPY php/ /opt/easy-excel/php/
LABEL org.opencontainers.image.title="frankenphp-easy-excel" \
      org.opencontainers.image.description="FrankenPHP with the easy-excel (Go/excelize) spreadsheet extension" \
      org.opencontainers.image.source="https://github.com/xiidea/easy-excel"

RUN apt-get update && apt-get install -y --no-install-recommends \
	acl \
	file \
	gettext \
	&& rm -rf /var/lib/apt/lists/*

RUN set -eux; \
	install-php-extensions \
		redis \
        mongodb \
        pdo_pgsql \
        pdo_mysql \
        bcmath \
            apcu \
            gd \
            intl \
            opcache \
            zip \
            igbinary \
            yaml \
            xsl \
            gettext \
            sockets \
	;
