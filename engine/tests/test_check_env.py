"""The launcher's setup check must match requirements.txt exactly.

A package listed in check_env.py but not installed by requirements.txt makes the
one-click setup fail after a successful install, so keep the two in sync.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path
import re
import unittest

ENGINE = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("check_env", ENGINE / "check_env.py")
check_env = importlib.util.module_from_spec(spec)
spec.loader.exec_module(check_env)

# Installed automatically as dependencies of a listed requirement.
DEPENDENCIES = {"lxml"}  # python-pptx


def requirement_names():
    names = set()
    for line in (ENGINE / "requirements.txt").read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if line and not line.startswith("-"):
            names.add(re.split(r"[<>=!~;\[ ]", line, maxsplit=1)[0].strip().lower())
    return names


class CheckEnvMatchesRequirements(unittest.TestCase):
    def test_every_checked_package_is_installed_by_requirements(self):
        checked = {dist.lower() for _, dist in check_env.REQUIRED}
        self.assertEqual(checked - DEPENDENCIES, requirement_names())


if __name__ == "__main__":
    unittest.main()
