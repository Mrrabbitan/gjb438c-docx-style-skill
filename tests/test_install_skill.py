"""Exercise replacement safety without touching the real installed skill."""

from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock


MODULE_PATH = Path(__file__).resolve().parents[1] / "tools" / "install_skill.py"
SPEC = importlib.util.spec_from_file_location("install_skill", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
installer = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(installer)


class InstallSkillTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.repo = self.root / "repo"
        self.source = self.repo / installer.SKILL_NAME
        self.source.mkdir(parents=True)
        (self.source / "SKILL.md").write_text("new skill", encoding="utf-8")
        self.patch = mock.patch.object(installer, "repo_root", return_value=self.repo)
        self.patch.start()
        self.addCleanup(self.patch.stop)
        self.dest = self.root / "skills"

    def create_installed(self):
        target = self.dest / installer.SKILL_NAME
        target.mkdir(parents=True)
        (target / "SKILL.md").write_text("installed version", encoding="utf-8")
        return target

    def make_symlink(self, link, destination):
        try:
            link.symlink_to(destination, target_is_directory=True)
        except OSError as error:
            self.skipTest(f"Directory symlinks unavailable: {error}")

    def test_copy_excludes_machine_caches(self):
        (self.source / ".DS_Store").write_bytes(b"metadata")
        cache = self.source / "__pycache__"
        cache.mkdir()
        (cache / "script.pyc").write_bytes(b"cache")
        target = installer.install("copy", self.dest, False)
        self.assertEqual((target / "SKILL.md").read_text(), "new skill")
        self.assertFalse((target / ".DS_Store").exists())
        self.assertFalse((target / "__pycache__").exists())

    def test_existing_install_requires_force(self):
        target = self.create_installed()
        with self.assertRaises(FileExistsError):
            installer.install("copy", self.dest, False)
        self.assertEqual((target / "SKILL.md").read_text(), "installed version")

    def test_force_keeps_previous_version(self):
        target = self.create_installed()
        installer.install("copy", target, True)
        backups = list(self.dest.glob(f"{installer.SKILL_NAME}.backup-*"))
        self.assertEqual(len(backups), 1)
        self.assertEqual((backups[0] / "SKILL.md").read_text(), "installed version")
        self.assertEqual((target / "SKILL.md").read_text(), "new skill")

    def test_exact_target_symlink_does_not_delete_source(self):
        self.dest.mkdir()
        target = self.dest / installer.SKILL_NAME
        self.make_symlink(target, self.source)
        installer.install("copy", target, True)
        self.assertFalse(target.is_symlink())
        self.assertEqual((self.source / "SKILL.md").read_text(), "new skill")
        self.assertEqual(len(list(self.dest.glob(f"{installer.SKILL_NAME}.backup-*"))), 1)

    def test_rejects_same_and_nested_paths(self):
        for dest in [self.source, self.source / "inside"]:
            with self.subTest(dest=dest), self.assertRaises(ValueError):
                installer.install("copy", dest, True)
        with self.assertRaises(ValueError):
            installer.check_paths(self.source, self.repo)
        self.assertEqual((self.source / "SKILL.md").read_text(), "new skill")

    def test_symlink_parent_cannot_hide_nested_target(self):
        alias = self.root / "alias"
        self.make_symlink(alias, self.source)
        with self.assertRaises(ValueError):
            installer.install("copy", alias, True)

    def test_copy_failure_preserves_installed_version(self):
        target = self.create_installed()
        with mock.patch.object(installer.shutil, "copytree", side_effect=OSError("copy failed")):
            with self.assertRaises(OSError):
                installer.install("copy", self.dest, True)
        self.assertEqual((target / "SKILL.md").read_text(), "installed version")
        self.assertFalse(list(self.dest.glob(".*-install-*")))

    def test_incomplete_staging_preserves_installed_version(self):
        target = self.create_installed()
        real_copy = installer.shutil.copytree

        def corrupt_staging(source, destination, **kwargs):
            result = real_copy(source, destination, **kwargs)
            (destination / "SKILL.md").write_text("incomplete copy", encoding="utf-8")
            return result

        with mock.patch.object(installer.shutil, "copytree", side_effect=corrupt_staging):
            with self.assertRaises(ValueError):
                installer.install("copy", self.dest, True)
        self.assertEqual((target / "SKILL.md").read_text(), "installed version")
        self.assertFalse(list(self.dest.glob(f"{installer.SKILL_NAME}.backup-*")))

    def test_replacement_failure_restores_installed_version(self):
        target = self.create_installed()
        real_rename = Path.rename

        def fail_staged_rename(path, destination):
            if path.parent.name.startswith(f".{installer.SKILL_NAME}-install-"):
                raise OSError("replace failed")
            return real_rename(path, destination)

        with mock.patch.object(Path, "rename", new=fail_staged_rename):
            with self.assertRaises(OSError):
                installer.install("copy", self.dest, True)
        self.assertEqual((target / "SKILL.md").read_text(), "installed version")
        self.assertFalse(list(self.dest.glob(".*-install-*")))

    def test_rejects_unbundled_symlink_before_replacement(self):
        target = self.create_installed()
        self.make_symlink(self.source / "external", self.dest)
        with self.assertRaises(ValueError):
            installer.install("copy", self.dest, True)
        self.assertEqual((target / "SKILL.md").read_text(), "installed version")

    def test_symlink_install_and_broken_link_replacement(self):
        self.dest.mkdir()
        target = self.dest / installer.SKILL_NAME
        self.make_symlink(target, self.root / "missing")
        installer.install("symlink", target, True)
        self.assertTrue(target.is_symlink())
        self.assertEqual(target.resolve(), self.source.resolve())


if __name__ == "__main__":
    unittest.main()
