#!/bin/sh
set -eu

cd "$(dirname "$0")/../../SpreadSheet.Tests"
framework="net${1:?Specify the .NET version, for example 10.0}"

dotnet --info
dotnet restore -p:TargetFrameworks="$framework"
# Restore library references separately so tests exercise the shipped .NET Standard assemblies.
dotnet restore ../SpreadSheet/SpreadSheet.csproj
dotnet test --configuration Release --no-restore --framework "$framework" --logger "trx;LogFileName=tests.trx"
