#!/bin/sh
# Builds the linux-musl native library and its musllinux wheel in Alpine containers, running the
# Python tests against it. Needs a Linux host with Docker; the wheel lands in dist/.
#
# Run from the repository root:  sh python/scripts/build_musl_wheel.sh
set -eu

docker run --rm -v "$PWD:/repo" -w /repo mcr.microsoft.com/dotnet/sdk:10.0-alpine \
  sh -c 'apk add --no-cache clang build-base zlib-dev python3 && python3 python/scripts/build_native.py'

docker run --rm -v "$PWD:/repo" -w /repo python:3.12-alpine \
  sh -c 'pip install -e "python[dev]" build && pytest python/tests && python -m build --wheel python/ --outdir dist/ && python python/scripts/check_wheel_tag.py dist'
