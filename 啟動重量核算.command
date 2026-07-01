#!/bin/bash
cd "$(dirname "$0")"
open http://localhost:8770
python3 -m http.server 8770
