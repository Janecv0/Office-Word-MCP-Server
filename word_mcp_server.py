#!/usr/bin/env python3
"""
Run script for the Word Document Server.

This script provides a simple way to start the Word Document Server.
"""

import os
from pathlib import Path

from word_document_server.main import run_server

# Output directory for saved documents
OUTPUT_DIR = Path(os.environ.get("MCP_OUTPUT_DIR", os.environ.get("PPT_OUTPUT_DIR", "./output"))).resolve()
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
os.environ["MCP_OUTPUT_DIR"] = str(OUTPUT_DIR)

if __name__ == "__main__":
    run_server()
