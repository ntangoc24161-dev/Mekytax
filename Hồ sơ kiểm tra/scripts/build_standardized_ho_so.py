# -*- coding: utf-8 -*-
"""
Main entry point for building and standardizing the complete inspection dossier:
- Pure template loading from MAUTONG HOP.docx
- Accurate placeholder substitution
- Dedicated Decree 30 formatting for Administrative Documents only
- 100% template preservation for Professional Inspection Forms
"""
from build_bien_ban_ho_so_ve_vpt import run_rebuild

if __name__ == "__main__":
    run_rebuild()
