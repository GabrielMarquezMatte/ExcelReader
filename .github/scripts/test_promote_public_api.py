#!/usr/bin/env python3
"""Unit tests for promote_public_api.py.

Run: python3 .github/scripts/test_promote_public_api.py
"""

import os
import tempfile
import unittest

import promote_public_api as promote


class PromoteTests(unittest.TestCase):
    def _write(self, path, lines):
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(promote.HEADER + "\n")
            for line in lines:
                f.write(line + "\n")

    def _read(self, path):
        with open(path, encoding="utf-8") as f:
            return {line.strip() for line in f if line.strip() and line.strip() != promote.HEADER}

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.shipped = os.path.join(self._tmp.name, "PublicAPI.Shipped.txt")
        self.unshipped = os.path.join(self._tmp.name, "PublicAPI.Unshipped.txt")

    def tearDown(self):
        self._tmp.cleanup()

    def test_plain_entry_is_added_to_shipped(self):
        self._write(self.shipped, ["Foo.Bar() -> void"])
        self._write(self.unshipped, ["Foo.Baz() -> void"])

        promote.promote_one(self.unshipped, self.shipped)

        self.assertEqual(self._read(self.shipped), {"Foo.Bar() -> void", "Foo.Baz() -> void"})
        self.assertEqual(self._read(self.unshipped), set())

    def test_removed_entry_deletes_the_plain_entry_from_shipped(self):
        # This is the exact regression this PR fixes: a *REMOVED* entry used to be unioned
        # into Shipped as a literal string instead of deleting its plain counterpart, leaving
        # the plain entry (a symbol no longer in the code) permanently stuck in Shipped.txt —
        # which the PublicApiAnalyzer then flags as RS0017 on every subsequent build.
        self._write(self.shipped, ["Foo.Bar() -> void", "Foo.Baz() -> void"])
        self._write(self.unshipped, ["*REMOVED*Foo.Bar() -> void"])

        promote.promote_one(self.unshipped, self.shipped)

        self.assertEqual(self._read(self.shipped), {"Foo.Baz() -> void"})
        self.assertEqual(self._read(self.unshipped), set())

    def test_removed_entry_is_never_written_to_shipped(self):
        # RS0024: "The shipped API file can't have removed members" - a *REMOVED* line is only
        # ever meaningful in Unshipped.txt (an instruction to delete on promotion); it must not
        # survive into Shipped.txt itself.
        self._write(self.shipped, ["Foo.Bar() -> void"])
        self._write(self.unshipped, ["*REMOVED*Foo.Bar() -> void"])

        promote.promote_one(self.unshipped, self.shipped)

        for entry in self._read(self.shipped):
            self.assertFalse(entry.startswith("*REMOVED*"), entry)

    def test_removed_entry_with_no_matching_shipped_entry_is_a_noop_delete(self):
        # Removing something that was added and removed within the same unshipped window (never
        # actually promoted as a plain entry) must not crash and must not leave any trace.
        self._write(self.shipped, ["Foo.Bar() -> void"])
        self._write(self.unshipped, ["*REMOVED*Foo.NeverShipped() -> void"])

        promote.promote_one(self.unshipped, self.shipped)

        self.assertEqual(self._read(self.shipped), {"Foo.Bar() -> void"})

    def test_mixed_add_and_remove_in_the_same_promotion(self):
        self._write(self.shipped, ["Foo.Old() -> void"])
        self._write(self.unshipped, ["*REMOVED*Foo.Old() -> void", "Foo.New() -> void"])

        promote.promote_one(self.unshipped, self.shipped)

        self.assertEqual(self._read(self.shipped), {"Foo.New() -> void"})


if __name__ == "__main__":
    unittest.main()
