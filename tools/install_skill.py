#!/usr/bin/env python3
"""Install the bundled Codex skill into a local skills directory."""

from __future__ import annotations

import argparse
import os
import shutil
from pathlib import Path


SKILL_NAME = "gjb438c-docx-style"


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def default_skills_root() -> Path:
    codex_home = os.environ.get("CODEX_HOME")
    if codex_home:
        return Path(codex_home).expanduser() / "skills"
    return Path.home() / ".codex" / "skills"


def resolve_target(dest: Path) -> Path:
    dest = dest.expanduser().resolve()
    if dest.name == SKILL_NAME:
        return dest
    return dest / SKILL_NAME


def remove_existing(path: Path) -> None:
    if path.is_symlink() or path.is_file():
        path.unlink()
    elif path.exists():
        shutil.rmtree(path)


def install(mode: str, dest: Path, force: bool) -> Path:
    source = repo_root() / SKILL_NAME
    if not source.exists():
        raise FileNotFoundError(f"Skill source directory not found: {source}")

    target = resolve_target(dest)
    target.parent.mkdir(parents=True, exist_ok=True)

    if target.exists() or target.is_symlink():
        if not force:
            raise FileExistsError(f"Target already exists: {target}. Use --force to replace it.")
        remove_existing(target)

    if mode == "copy":
        shutil.copytree(source, target)
    elif mode == "symlink":
        target.symlink_to(source, target_is_directory=True)
    else:
        raise ValueError(f"Unsupported install mode: {mode}")

    return target


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dest", type=Path, default=default_skills_root(), help="Codex skills root or exact skill target path.")
    parser.add_argument("--mode", choices=["copy", "symlink"], default="copy", help="Install by copying or symlinking the skill.")
    parser.add_argument("--force", action="store_true", help="Replace an existing target.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    target = install(args.mode, args.dest, args.force)
    print(f"Installed {SKILL_NAME} to {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

