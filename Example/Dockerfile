# syntax=docker/dockerfile:1.20

FROM mcr.microsoft.com/dotnet/core/sdk:2.2-alpine AS build

WORKDIR /src

COPY --parents Example/ SpreadSheet/ CompoundFile/ ./

RUN dotnet restore Example/Example.csproj -r linux-musl-x64 \
    && dotnet restore SpreadSheet/SpreadSheet.csproj -r linux-musl-x64

RUN dotnet publish Example/Example.csproj -c Release -o /app -r linux-musl-x64 \
    --self-contained false --no-restore

FROM mcr.microsoft.com/dotnet/core/runtime:2.2-alpine AS final

ARG TZ='Europe/Riga'

ENV DEFAULT_TZ=${TZ} \
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false \
    LOGGING__CONSOLE__DISABLECOLORS=true \
    LC_ALL=lv_LV.UTF-8 \
    LANG=lv_LV.UTF-8

RUN apk add --no-cache tzdata ca-certificates icu-libs \
    && cp /usr/share/zoneinfo/${DEFAULT_TZ} /etc/localtime \
    && addgroup -S app \
    && adduser -S -G app app \
    && mkdir -p /app \
    && chown app:app /app

WORKDIR /app
COPY --from=build --chown=app:app /app ./

USER app

ENTRYPOINT ["dotnet", "Example.dll"]
