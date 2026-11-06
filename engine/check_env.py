"""Environment self-check used by launcher.ps1.

Imports every required package and prints ONE line starting with ENVCHECK
followed by JSON, so the launcher can show exactly which package failed and why.
Exit code 0 = everything imports, 1 = something is missing or broken.
"""
import importlib
import json
import os
import sys
from importlib import metadata

# (import name, pip distribution name)
REQUIRED = [
    ("pptx", "python-pptx"),
    ("lxml", "lxml"),
    ("pypdf", "pypdf"),
    ("pypdfium2", "pypdfium2"),
    ("PIL", "Pillow"),
]


def main():
    prefix = os.path.normcase(os.path.realpath(sys.prefix))
    result = {"python": sys.version.split()[0], "executable": sys.executable,
              "ok": True, "packages": []}
    for module, dist in REQUIRED:
        entry = {"module": module, "package": dist, "version": None, "ok": False,
                 "error": None, "location": None, "outside_environment": False}
        try:
            entry["version"] = metadata.version(dist)
        except metadata.PackageNotFoundError:
            entry["error"] = "not installed"
        try:
            loaded = importlib.import_module(module)
            entry["ok"] = True
            entry["error"] = None
            location = getattr(loaded, "__file__", None)
            if location:
                entry["location"] = location
                real = os.path.normcase(os.path.realpath(location))
                entry["outside_environment"] = not real.startswith(prefix)
        except Exception as exc:  # report every failure, never crash the check
            entry["error"] = ("%s: %s" % (type(exc).__name__, exc)).strip()[:800]
            result["ok"] = False
        result["packages"].append(entry)
    print("ENVCHECK " + json.dumps(result))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
