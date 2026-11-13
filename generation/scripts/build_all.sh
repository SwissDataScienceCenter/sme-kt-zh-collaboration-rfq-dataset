#!/usr/bin/env bash
# Rebuild the whole Company world from the sources in generation/world/.
set -euo pipefail
cd "$(dirname "$0")/../.."
python3 generation/scripts/build_item_master.py
python3 generation/scripts/pricing.py
python3 generation/scripts/build_systems.py
python3 generation/scripts/render_organisation.py
python3 generation/scripts/render_operations.py
python3 generation/scripts/render_documents.py
python3 generation/scripts/render_regulations.py
