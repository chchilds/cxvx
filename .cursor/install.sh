#!/usr/bin/env bash
#
# Idempotent bootstrap for the actively developed component of this repo:
# the Cisco Enterprise Agreement PowerPoint generator under pptx/.
#
# The legacy Spark Operations Bot (app.py) pins 2017-era packages
# (Flask 0.12.2, gevent 1.2.2, MarkupSafe 1.0, ciscosparkbot) that require
# Python 3.6 and live Cisco Spark/Meraki/Umbrella credentials, so it is not
# provisioned here. Run it separately via its Dockerfile (FROM python:3.6).
set -euo pipefail

cd "$(dirname "$0")/.."

# The default image ships Python 3.12 but may lack the venv/ensurepip module.
if ! python3 -c 'import ensurepip' >/dev/null 2>&1; then
  sudo apt-get update -qq
  sudo apt-get install -y -qq python3-venv
fi

# Create or refresh an isolated virtualenv for the generator.
python3 -m venv .venv
# shellcheck disable=SC1091
source .venv/bin/activate

python -m pip install --upgrade pip
pip install -r pptx/requirements.txt

echo "Environment ready."
echo "Activate with: source .venv/bin/activate"
echo "Generate a deck: cd pptx && python generate_ea_summary.py <EAMP.xlsx> -o out.pptx [--dark]"
