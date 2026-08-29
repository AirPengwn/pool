#!/usr/bin/env python3
"""install_hooks.py -- copy tools/pre-commit into .git/hooks/.

Git hooks live in .git/, which is not version-controlled, so a hook committed
to the repo does nothing until it is installed. Run this once per clone:

    python tools/install_hooks.py

Idempotent: re-running overwrites the installed copy, so it also serves as the
way to pick up changes to tools/pre-commit.
"""
import os
import shutil
import stat
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
SRC = os.path.join(HERE, "pre-commit")
HOOKS = os.path.join(REPO, ".git", "hooks")
DST = os.path.join(HOOKS, "pre-commit")

if not os.path.isdir(HOOKS):
    sys.exit(f"no {HOOKS} -- is this a git clone?")
shutil.copyfile(SRC, DST)
os.chmod(DST, os.stat(DST).st_mode | stat.S_IEXEC | stat.S_IXGRP | stat.S_IXOTH)
print(f"installed {DST}")
print("verify with: git commit (it will run check_log.py first)")
