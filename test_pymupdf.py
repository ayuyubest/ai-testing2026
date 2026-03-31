#!/usr/bin/env python3
"""Test if PyMuPDF can be imported"""

try:
    print("Testing PyMuPDF import...")
    import fitz  # PyMuPDF
    print(f"✅ PyMuPDF imported successfully! Version: {fitz.version}")
except ImportError as e:
    print(f"❌ Import error: {e}")
    print("\nSolution: Install Visual C++ Redistributable")
    print("Download: https://aka.ms/vs/17/release/vc_redist.x64.exe")
except Exception as e:
    print(f"❌ DLL load error: {e}")
    print("\nThis is a Windows DLL issue. Solutions:")
    print("1. Install Visual C++ Redistributable: https://aka.ms/vs/17/release/vc_redist.x64.exe")
    print("2. Or reinstall PyMuPDF: pip uninstall pymupdf && pip install pymupdf")

