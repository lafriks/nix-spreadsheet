#!/bin/sh
set -eu

cd "$(dirname "$0")/../../SpreadSheet.Tests"
target="${1:?Specify the .NET version or framework, for example 10.0 or net48}"
case "$target" in
  net*) framework="$target" ;;
  *) framework="net$target" ;;
esac

dotnet --info
dotnet restore -p:TargetFramework="$framework" -p:TargetFrameworks="$framework"
# Restore library references separately so tests exercise the shipped .NET Standard assemblies.
dotnet restore ../SpreadSheet/SpreadSheet.csproj
case "$framework" in
  # Framework's default C# 7.3 cannot compile the existing nullable annotations and using declarations.
  net4*) set -- -p:LangVersion=8.0 ;;
  *) set -- ;;
esac
dotnet test --configuration Release --no-restore --framework "$framework" "$@" --logger "trx;LogFileName=tests.trx"
