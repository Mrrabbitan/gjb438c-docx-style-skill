#!/usr/bin/env python3
"""Install the bundled Codex skill into a local skills directory."""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path


SKILL_NAME = "gjb438c-docx-style"
IGNORE_PATTERNS = ("__pycache__", "*.pyc", "*.pyo", ".DS_Store", ".git", "~$*")


def repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def default_skills_root() -> Path:
    codex_home = os.environ.get("CODEX_HOME")
    if codex_home:
        return Path(codex_home).expanduser() / "skills"
    return Path.home() / ".codex" / "skills"


def resolve_target(dest: Path) -> Path:
    # Resolve parents, but never follow the leaf symlink we may replace.
    dest = Path(os.path.abspath(dest.expanduser()))
    candidate = dest if dest.name == SKILL_NAME else dest / SKILL_NAME
    return candidate.parent.resolve() / candidate.name


def remove_existing(path: Path) -> None:
    if path.is_symlink() or path.is_file():
        path.unlink()
    elif path.exists():
        shutil.rmtree(path)


def check_paths(source: Path, target: Path) -> None:
    if source == target or source in target.parents or target in source.parents:
        raise ValueError("Source and target must be separate, non-nested paths")


def bundle_manifest(path: Path) -> dict[str, str]:
    """Hash installable files and reject links to unbundled local data."""
    if not (path / "SKILL.md").is_file():
        raise ValueError(f"Not a skill bundle: {path}")
    manifest: dict[str, str] = {}
    ignore = shutil.ignore_patterns(*IGNORE_PATTERNS)
    for root, directories, files in os.walk(path):
        excluded = ignore(root, directories + files)
        directories[:] = sorted(name for name in directories if name not in excluded)
        for name in directories + sorted(name for name in files if name not in excluded):
            entry = Path(root) / name
            if entry.is_symlink():
                raise ValueError(f"Skill bundle must not contain symlinks: {entry}")
            if entry.is_file():
                manifest[str(entry.relative_to(path))] = hashlib.sha256(entry.read_bytes()).hexdigest()
    return manifest


def backup_path(target: Path) -> Path:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return target.with_name(f"{target.name}.backup-{stamp}-{uuid.uuid4().hex[:8]}")


def install(mode: str, dest: Path, force: bool) -> Path:
    if mode not in ("copy", "symlink"):
        raise ValueError(f"Unsupported install mode: {mode}")
    source = (repo_root() / SKILL_NAME).resolve()
    if not source.is_dir():
        raise FileNotFoundError(f"Skill source directory not found: {source}")

    target = resolve_target(dest)
    check_paths(source, target)
    if target.exists() or target.is_symlink():
        if not force:
            raise FileExistsError(f"Target already exists: {target}. Use --force to replace it.")

    source_manifest = bundle_manifest(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    staging_root = Path(tempfile.mkdtemp(prefix=f".{SKILL_NAME}-install-", dir=target.parent))
    staged = staging_root / SKILL_NAME
    backup: Path | None = None
    try:
        if mode == "copy":
            shutil.copytree(source, staged, ignore=shutil.ignore_patterns(*IGNORE_PATTERNS))
            if bundle_manifest(staged) != source_manifest:
                raise ValueError("Staged skill does not match the source bundle")
        else:
            staged.symlink_to(source, target_is_directory=True)
            if not staged.is_dir() or bundle_manifest(staged) != source_manifest:
                raise ValueError("Staged skill symlink does not resolve to the source bundle")
        # Detect a concurrent source edit before replacing an installed version.
        if bundle_manifest(source) != source_manifest:
            raise ValueError("Skill source changed during installation; retry")
        if target.exists() or target.is_symlink():
            if not force:
                raise FileExistsError(f"Target appeared during installation: {target}")
            backup = backup_path(target)
            target.rename(backup)
        try:
            staged.rename(target)
        except BaseException:
            if backup is not None and not target.exists() and not target.is_symlink():
                backup.rename(target)
            raise
    finally:
        remove_existing(staging_root)
    if backup is not None:
        print(f"Previous installation preserved at {backup}")
    return target


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dest", type=Path, default=default_skills_root(), help="Codex skills root or exact skill target path.")
    parser.add_argument("--mode", choices=["copy", "symlink"], default="copy", help="Install by copying or symlinking the skill.")
    parser.add_argument("--force", action="store_true", help="Replace an existing target after validation; preserve it as a sibling backup.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    target = install(args.mode, args.dest, args.force)
    print(f"Installed {SKILL_NAME} to {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
