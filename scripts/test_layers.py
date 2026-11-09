"""Versioned module classification and unique test outcome accounting (stdlib only)."""

import fnmatch
import importlib.abc
import json
from pathlib import Path
import sys
import unittest

LAYERS = ("unit", "integration", "qualification")
COUNTS = ("passed", "failed", "skipped", "pending")


def unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("Duplicate JSON key")
        result[key] = value
    return result


def load_classification(path, tests_dir):
    data = json.loads(Path(path).read_text(encoding="utf-8"), object_pairs_hook=unique_object)
    if (not isinstance(data, dict) or data.get("schema_version") != "1.0"
            or not isinstance(data.get("inventory_version"), str) or not data["inventory_version"]):
        raise ValueError("Unsupported classification version")
    layers = data.get("layers")
    if not isinstance(layers, dict) or set(layers) != {"unit", "integration"}:
        raise ValueError("Classification must assign unit and integration modules")
    assignments = {}
    for layer, names in layers.items():
        if not isinstance(names, list):
            raise ValueError("Classification module list required")
        for name in names:
            if (not isinstance(name, str) or Path(name).name != name
                    or not name.startswith("test_") or not name.endswith(".py") or name in assignments):
                raise ValueError("Invalid or duplicate classification")
            assignments[name] = layer
    inventory = {path.name for path in Path(tests_dir).glob("test_*.py") if path.is_file()}
    if inventory != set(assignments):
        missing = sorted(inventory - set(assignments))
        stale = sorted(set(assignments) - inventory)
        raise ValueError(f"Classification mismatch: missing={missing}, stale={stale}")
    return assignments


class UnitImportGuard(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname.split(".")[0] in {"fitz", "pymupdf", "PySide6", "PyQt5", "PyQt6"}:
            raise RuntimeError("Unit layer cannot import GUI/native PDF engines")
        return None


def install_unit_guard():
    sys.meta_path.insert(0, UnitImportGuard())

    def no_external_process(event, args):
        if event in {"subprocess.Popen", "os.system", "os.posix_spawn", "os.spawn", "os.exec", "os.fork"}:
            raise RuntimeError("Unit layer cannot start external processes")

    sys.addaudithook(no_external_process)


def flatten(suite):
    for test in suite:
        if isinstance(test, unittest.TestSuite):
            yield from flatten(test)
        else:
            yield test


def select_suite(tests_dir, assignments, pattern, selected):
    suite = unittest.TestSuite()
    test_layers = {}
    loader = unittest.TestLoader()
    for filename, layer in sorted(assignments.items()):
        if layer not in selected or not fnmatch.fnmatchcase(filename, pattern):
            continue
        discovered = loader.discover(str(tests_dir), pattern=filename)
        for test in flatten(discovered):
            test_id = test.id()
            if test_id in test_layers:
                raise ValueError("Duplicate discovered test ID")
            test_layers[test_id] = layer
            suite.addTest(test)
    if not test_layers:
        raise ValueError("No tests selected")
    return suite, test_layers


class LayerResult(unittest.TextTestResult):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.outcomes = {}

    def startTest(self, test):
        self.outcomes[test.id()] = "passed"
        super().startTest(test)

    def addFailure(self, test, err):
        self.outcomes[test.id()] = "failed"
        super().addFailure(test, err)

    def addError(self, test, err):
        self.outcomes[test.id()] = "failed"
        super().addError(test, err)

    def addSubTest(self, test, subtest, err):
        if err is not None:
            self.outcomes[test.id()] = "failed"
        super().addSubTest(test, subtest, err)

    def addSkip(self, test, reason):
        # A skipped subtest must not erase another subtest's failure.
        parent = getattr(test, "test_case", test)
        if self.outcomes.get(parent.id()) != "failed":
            self.outcomes[parent.id()] = "skipped"
        super().addSkip(test, reason)

    def addExpectedFailure(self, test, err):
        self.outcomes[test.id()] = "skipped"
        super().addExpectedFailure(test, err)

    def addUnexpectedSuccess(self, test):
        self.outcomes[test.id()] = "failed"
        super().addUnexpectedSuccess(test)


def empty_summary():
    return {"schema_version": "1.0", "layers": {
        layer: {"status": "not_selected", "count_basis": "declared_sample" if layer == "qualification" else "unittest_test_case",
            **dict.fromkeys(COUNTS, 0)} for layer in LAYERS},
        "totals": dict.fromkeys(COUNTS, 0)}


def summarize_tests(summary, result, test_layers, selected):
    outcomes = {test_id: result.outcomes.get(test_id, "pending") for test_id in test_layers}
    suite_errors = dict.fromkeys(selected, 0)
    # unittest reports class/module fixture failures and skips as _ErrorHolder IDs.
    # They are suite outcomes, not extra tests and not successful unrun methods.
    for group_id, outcome in result.outcomes.items():
        if group_id in test_layers:
            continue
        prefix = group_id.partition("(")[2].removesuffix(")")
        affected = [test_id for test_id in test_layers if prefix and test_id.startswith(prefix + ".")]
        if outcome == "skipped":
            for test_id in affected:
                if outcomes[test_id] == "pending":
                    outcomes[test_id] = "skipped"
        elif outcome == "failed":
            for layer in {test_layers[test_id] for test_id in affected} or set(selected):
                suite_errors[layer] += 1
    for layer in selected:
        item = summary["layers"][layer]
        values = [outcome for test_id, outcome in outcomes.items() if test_layers[test_id] == layer]
        for status in COUNTS:
            item[status] = values.count(status)
        item["suite_errors"] = suite_errors[layer]
        item["status"] = "failed" if item["failed"] or item["suite_errors"] else "pending" if item["pending"] else "passed" if item["passed"] else "skipped" if item["skipped"] else "not_selected"
    return finish_summary(summary)


def finish_summary(summary):
    summary["totals"] = {status: sum(layer[status] for layer in summary["layers"].values()) for status in COUNTS}
    return summary
