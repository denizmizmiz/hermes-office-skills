#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""Shared contract layer for the xlsx skill scripts.

This is the SINGLE place where the Phase-1 safety invariants are enforced:

  I2  data_only guard   -- a workbook opened with data_only=True can never
                           be saved (formulas would silently become values).
  I1  read_only split   -- read_only workbooks are analysis-only; they are
                           rejected by the write path.
  I3  no silent skip    -- every result carries warnings[]/unsupported[]/
                           diagnostics[]; unknown spec keys are never dropped
                           silently on a write path.
  I5  safe write        -- backup -> temp output -> re-open + validate ->
                           atomic replace. Fail closed at every step.

Every other script imports this module and must never call `wb.save()`
directly (enforced by tests/test_xlsx_common.py::test_no_raw_save_outside_common).

Only stdlib + openpyxl are used here: this module must not import any other
skill module (one-way dependency; see SKILL.md architecture section).
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

# --------------------------------------------------------------------------
# dependency policy  (D1)
# --------------------------------------------------------------------------

REQUIREMENTS_PATH = Path(__file__).resolve().parent / "requirements.txt"
PIN_SPEC = "openpyxl>=3.1.5,<3.2"
SUPPORTED_OPENPYXL_MIN = (3, 1, 5)
SUPPORTED_OPENPYXL_MAX = (3, 2)


def _version_tuple(text: str) -> tuple:
    parts = []
    for chunk in str(text).split("."):
        digits = "".join(ch for ch in chunk if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def install_hint() -> str:
    """Exact commands that work on this host (measured)."""
    return (
        "uv run --with-requirements scripts/requirements.txt python "
        "scripts/<script>.py ...   # isolated, pinned, no environment change"
    )


def check_dependency() -> dict:
    """Verify openpyxl is importable and inside the pinned range. Read-only."""
    try:
        import openpyxl  # noqa: PLC0415
    except ImportError:
        return {
            "ok": False,
            "error_code": "DEPENDENCY_MISSING",
            "message": "openpyxl is not importable by the Python running this "
                       "script.",
            "recovery": (
                "Run the script through the pinned isolated environment: "
                + install_hint()
            ),
            "context": {
                "python": sys.executable,
                "python_version": sys.version.split()[0],
                "pin": PIN_SPEC,
                "requirements": str(REQUIREMENTS_PATH),
            },
        }
    version = getattr(openpyxl, "__version__", "unknown")
    vt = _version_tuple(version)
    if vt and (vt < SUPPORTED_OPENPYXL_MIN or vt >= SUPPORTED_OPENPYXL_MAX):
        return {
            "ok": False,
            "error_code": "DEPENDENCY_MISSING",
            "message": (f"openpyxl {version} is outside the pinned range "
                        f"{PIN_SPEC}."),
            "recovery": (
                "Install the pinned version, e.g. "
                "`pip install \"openpyxl>=3.1.5,<3.2\"`, or run through "
                + install_hint()
            ),
            "context": {"found": version, "pin": PIN_SPEC},
        }
    return {"ok": True, "openpyxl": version,
            "python": sys.version.split()[0], "python_executable": sys.executable}


def require_dependency() -> dict | None:
    """Return an error dict when the dependency is unusable, else None."""
    status = check_dependency()
    return None if status.get("ok") else status


# --------------------------------------------------------------------------
# error codes (closed set)
# --------------------------------------------------------------------------

ERROR_CODES = (
    "DATA_ONLY_WRITE_FORBIDDEN",
    "FILE_NOT_FOUND",
    "READ_ONLY_FILE",
    "DEPENDENCY_MISSING",
    "UNSUPPORTED_FEATURE",
    "SPEC_INVALID",
    "SPEC_UNKNOWN_KEY",
    "TABLE_NOT_FOUND",
    "SHEET_NOT_FOUND",
    "PLAN_CONFLICT",
    "BACKUP_FAILED",
    "WRITE_FAILED",
    "VALIDATION_FAILED",
    "QA_FAILED",
    # FAZ 3A execution codes (added with the template execution core; the
    # Phase-1 codes above are unchanged -- this only widens the closed set).
    "APPROVAL_REQUIRED",
    "APPROVAL_INVALID",
    "PLAN_INVALID",
    "STALE_PLAN",
    "EXECUTION_BLOCKED",
    "SOURCE_CHANGED",
    "TYPE_MISMATCH",
    "VALIDATION_MISMATCH",
    "MERGED_CELL_WRITE_FORBIDDEN",
    "FORMULA_MODIFICATION_FORBIDDEN",
    "ROW_EXPANSION_REQUIRED",
    "UNEXPECTED_WRITE",
    "UNEXPECTED_MODIFICATION",
    # not in the original minimum list: honest reporting for exceptions that
    # belong to no specific category (guard fallback, decode errors, ...).
    "UNEXPECTED_ERROR",
    # FAZ 3B row-expansion codes (widening only; FAZ 1/2/3A codes above
    # are unchanged).
    "PLAN_TRUNCATED",
    "EXPANSION_BOUNDARY_UNKNOWN",
    "MERGE_EXPANSION_UNSUPPORTED",
    "TABLE_EXPANSION_UNSUPPORTED",
    "EXPANSION_MULTI_BLOCK_UNSUPPORTED",
    "UNSUPPORTED_FORMULA_PROPAGATION",
    "RECORD_KEY_UNAVAILABLE",
    "DUPLICATE_RECORD",
    # FAZ 3C execution-hardening codes (widening only; FAZ 1/2/3A/3B codes
    # above are unchanged).
    "MANIFEST_MISSING",
    "MANIFEST_CORRUPT",
    "MANIFEST_LOCKED",
    "MANIFEST_WRITE_FAILED",
    "ATOMIC_COMMIT_FAILED",
    "IDEMPOTENCY_UNVERIFIED",
    # lookup taxonomy: xlsx_expand already raised these; registering them
    # stops XlsxError from silently rewriting them to WRITE_FAILED.
    "LOOKUP_NOT_FOUND",
    "LOOKUP_AMBIGUOUS",
    "LOOKUP_TABLE_NOT_FOUND",
    "LOOKUP_TABLE_EMPTY",
    "LOOKUP_KEY_COLUMN_MISSING",
    "LOOKUP_RESULT_COLUMN_MISSING",
    "LOOKUP_TYPE_MISMATCH",
    "UNSUPPORTED_LOOKUP",
    "INVALID_DV_RANGE",
)

WARNING_CODES = (
    "UNKNOWN_SPEC_KEY",
    "SILENT_NOOP_PREVENTED",
    "FEATURE_PARTIAL",
    "LOCALE_NORMALIZED",
    "RECALC_UNAVAILABLE",
    "READ_ONLY_NOOP",
    # FAZ 3B expansion warnings (widening only).
    "NO_EXPANSION_EVIDENCE",
    "ROW_EXPANSION_FLAG_RECONCILED",
    "AGGREGATE_MAY_NOT_COVER_NEW_ROWS",
    "PRESERVE_WITH_WARNING",
    "UNSUPPORTED_VALIDATION_PROPAGATION",
    # FAZ 3C hardening warnings (widening only).
    "MANIFEST_DUPLICATE",
    "MANIFEST_TRIM_REPORTED",
    "LOOKUP_RESULTS_NOT_MERGED",
)


class XlsxError(Exception):
    """Structured error carrying an error_code, recovery text and context."""

    def __init__(self, code: str, message: str, *, recovery=None, context=None):
        super().__init__(message)
        self.code = code if code in ERROR_CODES else "WRITE_FAILED"
        self.message = message
        self.recovery = recovery
        self.context = context or {}

    def as_dict(self) -> dict:
        payload = {
            "ok": False,
            # legacy alias: pre-Phase-1 parsers read the "error" field
            "error": self.message,
            "error_code": self.code,
            "message": self.message,
        }
        if self.recovery:
            payload["recovery"] = self.recovery
        if self.context:
            payload["context"] = self.context
        return payload


# --------------------------------------------------------------------------
# fault injection  (FAZ 3C, spec section 11)
# --------------------------------------------------------------------------

#: Deterministic injection registry: point name -> error code to raise.
#: Production runs leave this EMPTY (behaviour is then byte-identical to
#: the pre-3C code); tests set one entry, run, and clear it again.
FAULTS: dict = {}

#: The closed set of injection points the 3C suite exercises (10 failure
#: stages + the pre-commit validation stage).
FAULT_POINTS = (
    "backup",
    "temp_create",
    "staged_save",
    "reopen_validate",
    "pre_commit_validate",
    "expansion",
    "formula_propagation",
    "lookup",
    "qa",
    "manifest_append",
    "atomic_replace",
)


def fault(point: str, message: str, *, recovery=None, **context) -> None:
    """No-op unless a test registered ``FAULTS[point]``.

    Raising here is the ONLY effect of the registry: nothing is written,
    no state changes, and the error carries ``fault_point`` so a failure
    can never be mistaken for a production defect.
    """
    code = FAULTS.get(point)
    if not code:
        return
    raise XlsxError(
        code,
        f"injected fault at '{point}': {message}",
        recovery=recovery or ("Fault injection is a test-only switch "
                              "(xlsx_common.FAULTS); production never sets it."),
        context={"fault_point": point, **context},
    )


# --------------------------------------------------------------------------
# result envelope  (I3)
# --------------------------------------------------------------------------

class Result:
    """Accumulates one script run's outcome, then emits a single JSON object."""

    def __init__(self, **fields):
        self.payload = {"ok": True}
        self.payload.update(fields)
        self.warnings: list[dict] = []
        self.unsupported: list[dict] = []
        self.diagnostics: list[dict] = []

    def set(self, **fields) -> "Result":
        self.payload.update(fields)
        return self

    def get(self, key, default=None):
        return self.payload.get(key, default)

    def warn(self, code: str, message: str, **context) -> "Result":
        self.warnings.append({"code": code, "message": message,
                              **({"context": context} if context else {})})
        return self

    def unsupported_item(self, message: str, **context) -> "Result":
        self.unsupported.append({"message": message,
                                 **({"context": context} if context else {})})
        return self

    def diagnose(self, message: str, **context) -> "Result":
        self.diagnostics.append({"message": message,
                                 **({"context": context} if context else {})})
        return self

    def expect_change(self, changes, *, action: str) -> None:
        """I3: an operation was requested but nothing changed -> warn loudly."""
        if not changes:
            self.warn(
                "SILENT_NOOP_PREVENTED",
                f"'{action}' was requested but produced no change.",
                action=action,
            )

    def emit(self, *, exit_code: int = 0, stream=None) -> int:
        payload = dict(self.payload)
        payload["warnings"] = self.warnings
        payload["unsupported"] = self.unsupported
        payload["diagnostics"] = self.diagnostics
        print(json.dumps(payload, ensure_ascii=False, default=_json_default),
              file=stream or sys.stdout)
        return exit_code


def _json_default(value):
    from datetime import date, datetime, time  # noqa: PLC0415
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value)


def emit_result(payload: dict, *, warnings=None, unsupported=None,
                diagnostics=None, exit_code: int = 0) -> int:
    """Emit one result object; the three list fields are always present."""
    out = dict(payload)
    out.setdefault("ok", True)
    out["warnings"] = list(warnings or [])
    out["unsupported"] = list(unsupported or [])
    out["diagnostics"] = list(diagnostics or [])
    print(json.dumps(out, ensure_ascii=False, default=_json_default))
    return exit_code


def emit_error(error_code: str, message: str, *, recovery=None, context=None,
               extra=None) -> int:
    """Emit a structured failure to stderr and return exit code 1."""
    payload = {
        "ok": False,
        # legacy alias: pre-Phase-1 parsers read the "error" field
        "error": message,
        "error_code": error_code if error_code in ERROR_CODES
        else "UNEXPECTED_ERROR",
        "message": message,
    }
    if recovery:
        payload["recovery"] = recovery
    if context:
        payload["context"] = context
    if extra:
        payload.update(extra)
    print(json.dumps(payload, ensure_ascii=False, default=_json_default),
          file=sys.stderr)
    return 1


def dependency_error(exc: Exception = None) -> int:
    """Emit the structured DEPENDENCY_MISSING error (for import guards).

    Scripts with top-level openpyxl imports call this from an except
    ImportError block so a missing dependency still produces the structured
    contract instead of a raw traceback.
    """
    status = check_dependency()
    if status.get("ok"):
        return emit_error(
            "DEPENDENCY_MISSING",
            f"Required openpyxl component failed to import: {exc}",
            recovery=install_hint(),
            context={"exception": type(exc).__name__ if exc else None,
                     "python": sys.executable})
    return emit_error("DEPENDENCY_MISSING", status["message"],
                      recovery=status.get("recovery"),
                      context=status.get("context"))


def guard(main_fn):
    """Decorator: turn XlsxError and unexpected exceptions into JSON failures.

    Raw KeyError/ValueError/TypeError never reach the caller as a bare message;
    they are wrapped with the error code and (when available) the spec path.
    """
    import functools  # noqa: PLC0415

    @functools.wraps(main_fn)
    def wrapper(*args, **kwargs):
        dependency = require_dependency()
        if dependency is not None:
            return emit_error("DEPENDENCY_MISSING", dependency["message"],
                              recovery=dependency.get("recovery"),
                              context=dependency.get("context"))
        try:
            return main_fn(*args, **kwargs)
        except XlsxError as exc:
            print(json.dumps(exc.as_dict(), ensure_ascii=False,
                             default=_json_default), file=sys.stderr)
            return 1
        except KeyboardInterrupt:
            raise
        except Exception as exc:  # noqa: BLE001
            return emit_error(
                "UNEXPECTED_ERROR",
                f"{type(exc).__name__}: {exc}",
                recovery="Re-run with --help to check the arguments; if this "
                         "persists, the input file or spec is malformed.",
                context={"exception": type(exc).__name__},
            )

    return wrapper


# --------------------------------------------------------------------------
# shared value helpers  (canonical single copies; formerly duplicated in
# xlsx_edit.py / csv_to_xlsx.py (infer) and xlsx_read.py (jsonable))
# --------------------------------------------------------------------------

def jsonable(value):
    """JSON-safe scalar: datetimes/dates/times become ISO strings."""
    from datetime import date, datetime, time  # noqa: PLC0415
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return value


def infer(text, *, keep_empty_string: bool = False):
    """Type-infer CLI/CSV text: bool / int / float / ISO date / ISO datetime.

    Everything else -- including formulas -- stays a string. ``''`` maps to
    None for CSV semantics; pass keep_empty_string=True for the --set
    semantics of xlsx_edit.py, which has always written '' as an empty string.
    """
    if text == "":
        return "" if keep_empty_string else None
    if text.startswith("="):
        return text  # formula: openpyxl stores it as one
    low = text.lower()
    if low in ("true", "false"):
        return low == "true"
    from datetime import date, datetime  # noqa: PLC0415
    for caster in (int, float):
        try:
            return caster(text)
        except ValueError:
            pass
    for parser in (date.fromisoformat, datetime.fromisoformat):
        try:
            return parser(text)
        except ValueError:
            pass
    return text


# --------------------------------------------------------------------------
# spec validation  (D2/D3)
# --------------------------------------------------------------------------

class Schema:
    """Tiny declarative schema: leaf types, containers, required keys.

    Deliberately small: it validates the real spec shapes of xlsx_create.py and
    xlsx_edit.py and nothing more. No external schema library is introduced.
    """

    def __init__(self, *, required=(), optional=(), types=None, children=None,
                 many=False, typed_map=None):
        self.required = set(required)
        self.optional = set(optional)
        self.types = types or {}
        self.children = children or {}       # key -> Schema (nested object(s))
        self.many = set(many) if many else set()   # keys whose value is a list
        self.typed_map = typed_map or {}     # key -> Schema for dict valu es

    @property
    def known(self) -> set:
        return self.required | self.optional


def _describe(value) -> str:
    if value is None:
        return "null"
    return type(value).__name__


def _check_type(value, expected, path, errors, coerced):
    if expected == "str":
        ok = isinstance(value, str)
    elif expected == "bool":
        ok = isinstance(value, bool)
    elif expected == "number":
        ok = isinstance(value, (int, float)) and not isinstance(value, bool)
    elif expected == "int":
        ok = isinstance(value, int) and not isinstance(value, bool)
    elif expected == "dict":
        ok = isinstance(value, dict)
    elif expected == "list":
        ok = isinstance(value, list)
    elif expected == "str_or_dict":
        ok = isinstance(value, (str, dict))
    else:  # pragma: no cover - schema bug guard
        ok = True
    if not ok:
        errors.append({"path": path, "expected": expected,
                       "actual": _describe(value)})
    return ok


def validate_spec(spec, schema: Schema, *, path: str = "", strict: bool = False,
                  allow_unknown: bool = False) -> dict:
    """Validate a spec fragment against a Schema.

    Returns {"ok", "errors": [...], "unknown_keys": [...], "coerced": [...]}.
    In strict mode (write paths) an unknown key becomes a SPEC_UNKNOWN_KEY
    error unless allow_unknown is set; otherwise it lands in unknown_keys and
    the caller turns it into a warning (I3: never silently dropped).
    """
    errors: list[dict] = []
    unknown: list[dict] = []
    coerced: list[dict] = []

    if not isinstance(spec, dict):
        return {"ok": False,
                "errors": [{"path": path or "<root>", "expected": "dict",
                            "actual": _describe(spec)}],
                "unknown_keys": [], "coerced": []}

    for key in sorted(schema.required):
        if key not in spec:
            errors.append({"path": f"{path}.{key}" if path else key,
                           "expected": "required key present",
                           "actual": "missing"})

    for key, value in spec.items():
        here = f"{path}.{key}" if path else key
        if key not in schema.known:
            unknown.append({"path": here, "value_type": _describe(value)})
            continue
        expected = schema.types.get(key)
        if expected is not None:
            if not _check_type(value, expected, here, errors, coerced):
                continue
        child = schema.children.get(key)
        if child is None:
            continue
        if key in schema.many or isinstance(value, list):
            if not isinstance(value, list):
                continue
            for index, item in enumerate(value):
                sub = validate_spec(item, child, path=f"{here}[{index}]",
                                    strict=strict,
                                    allow_unknown=allow_unknown)
                errors.extend(sub["errors"])
                unknown.extend(sub["unknown_keys"])
                coerced.extend(sub["coerced"])
        elif isinstance(value, dict):
            sub = validate_spec(value, child, path=here, strict=strict,
                                allow_unknown=allow_unknown)
            errors.extend(sub["errors"])
            unknown.extend(sub["unknown_keys"])
            coerced.extend(sub["coerced"])

    if strict and unknown and not allow_unknown:
        for item in unknown:
            errors.append({
                "path": item["path"],
                "expected": "known key",
                "actual": f"unknown key ({item['value_type']})",
                "nearest": nearest_key(item["path"].split(".")[-1],
                                       schema.known),
            })

    return {"ok": not errors, "errors": errors, "unknown_keys": unknown,
            "coerced": coerced}


def nearest_key(name: str, candidates) -> str | None:
    """Cheap suggestion for a mistyped spec key (difflib, stdlib)."""
    import difflib  # noqa: PLC0415
    match = difflib.get_close_matches(name, sorted(candidates), n=1, cutoff=0.5)
    return match[0] if match else None


def spec_error(result: Result, validation: dict, *, strict: bool) -> bool:
    """Fold a validate_spec() outcome into a Result. True when fatal."""
    for item in validation.get("errors", []):
        if strict and item.get("expected") == "known key":
            continue  # recorded as SPEC_UNKNOWN_KEY below
        raise XlsxError(
            "SPEC_INVALID",
            f"Spec error at '{item['path']}': expected {item['expected']}, "
            f"got {item['actual']}.",
            recovery="Fix the spec field named in 'path' and re-run.",
            context={"path": item["path"], "expected": item["expected"],
                     "actual": item["actual"]},
        )
    unknown = validation.get("unknown_keys", [])
    if unknown:
        if strict:
            first = unknown[0]
            raise XlsxError(
                "SPEC_UNKNOWN_KEY",
                f"Unknown spec key '{first['path']}'.",
                recovery=("Remove the key, correct it, or pass "
                          "--allow-unknown-keys to accept it with a warning."),
                context={
                    "unknown_keys": [u["path"] for u in unknown],
                    "nearest": nearest_key(first["path"].split(".")[-1],
                                           _ALL_KNOWN_HINT),
                },
            )
        for item in unknown:
            result.warn("UNKNOWN_SPEC_KEY",
                        f"Unknown spec key '{item['path']}' was ignored.",
                        path=item["path"])
    return False


_ALL_KNOWN_HINT: set = set()


def register_known_keys(keys) -> None:
    """Let scripts advertise their full key set for 'nearest key' hints."""
    _ALL_KNOWN_HINT.update(keys)


# --------------------------------------------------------------------------
# safe loading  (I1 / I2)
# --------------------------------------------------------------------------

class LoadedWorkbook:
    """A workbook plus the authority it was loaded with.

    Modes:
      read_only    -- analysis only (read_only=True); the object has NO write
                      authority and carries no usable `save` path.
      values_only  -- cached formula values (data_only=True); `.wb` is not
                      exposed at all, so I2 cannot be bypassed structurally.
      normal       -- execution/modification; the only mode a save accepts.
    """

    def __init__(self, wb, *, path, mode, fingerprint=None):
        self._wb = wb
        self.path = str(path) if path else None
        self.mode = mode
        self.fingerprint = fingerprint
        self.data_only = mode == "values_only"
        self.read_only = mode == "read_only"

    # -- capability flags ---------------------------------------------------
    @property
    def writable(self) -> bool:
        return self.mode == "normal"

    @property
    def wb(self):
        """The live openpyxl workbook -- normal mode only (I2 structural guard)."""
        if self.mode == "values_only":
            raise XlsxError(
                "DATA_ONLY_WRITE_FORBIDDEN",
                "This workbook was opened with data_only=True; the live "
                "workbook object is deliberately not exposed, because saving "
                "it would replace every formula with its cached value.",
                recovery=("Open the file again with data_only=False for any "
                          "read/write work; keep this object for value reads."),
                context={"path": self.path, "mode": self.mode},
            )
        return self._wb

    # -- value access for values_only mode ----------------------------------
    @property
    def sheetnames(self):
        """Sheet names are always readable, even on values_only handles."""
        return list(self._wb.sheetnames)

    @property
    def active_sheetname(self):
        """Title of the sheet marked active (readable in every mode)."""
        return self._wb.active.title

    def rows(self, sheet):
        """All cell values of one sheet as JSON-ready rows, in every mode."""
        ws = self._wb[sheet]
        return [[jsonable(v) for v in row]
                for row in ws.iter_rows(values_only=True)]

    def values(self, sheet: str, ref: str = None):
        if self.mode != "values_only":
            raise XlsxError("READ_ONLY_FILE",
                            "values() is only valid on a values_only handle.",
                            context={"mode": self.mode})
        ws = self._wb[sheet]
        if ref is None:
            return [[c for c in row] for row in ws.iter_rows(values_only=True)]
        return ws[ref].value

    def iter_values(self, sheet: str):
        if self.mode != "values_only":
            raise XlsxError("READ_ONLY_FILE",
                            "iter_values() is only valid on a values_only "
                            "handle.", context={"mode": self.mode})
        return self._wb[sheet].iter_rows(values_only=True)

    def close(self) -> None:
        try:
            if self._wb is not None and self.mode == "read_only":
                self._wb.close()
        except Exception:  # noqa: BLE001
            pass


def load_workbook_safe(path, *, data_only=False, read_only=False,
                       keep_vba=None, keep_links=True) -> LoadedWorkbook:
    """The single loading entry point for the whole skill (I1/I2).

    Pass A  read_only=True,  data_only=False -> structure/formula text
    Pass B  read_only=True,  data_only=True  -> cached values only
    Pass C  read_only=False, data_only=False -> execution and saving
    The three objects are never handed to each other.
    """
    from openpyxl import load_workbook  # noqa: PLC0415

    target = Path(path)
    if not target.exists():
        raise XlsxError("FILE_NOT_FOUND", f"No such file: {target}",
                        recovery="Check the path and try again.",
                        context={"path": str(target)})
    if not target.is_file():
        raise XlsxError("READ_ONLY_FILE", f"Not a regular file: {target}",
                        context={"path": str(target)})

    kwargs = {"data_only": bool(data_only), "read_only": bool(read_only),
              "keep_links": bool(keep_links)}
    if keep_vba is not None:
        kwargs["keep_vba"] = bool(keep_vba)
    try:
        wb = load_workbook(str(target), **kwargs)
    except zipfile.BadZipFile as exc:
        raise XlsxError("VALIDATION_FAILED",
                        f"Not a readable xlsx package: {exc}",
                        recovery="The file is corrupt or is not an .xlsx "
                                 "workbook.",
                        context={"path": str(target)}) from exc
    except Exception as exc:  # noqa: BLE001
        raise XlsxError("VALIDATION_FAILED",
                        f"Could not open workbook: {type(exc).__name__}: {exc}",
                        context={"path": str(target)}) from exc

    # I2: data_only handles are structurally non-writable, regardless of the
    # read_only flag. read_only (analysis) and normal (execution) stay apart.
    if data_only:
        mode = "values_only"
    elif read_only:
        mode = "read_only"
    else:
        mode = "normal"
    return LoadedWorkbook(wb, path=target, mode=mode)


def fingerprint_file(path) -> str:
    """Deterministic file identity: size + mtime + sha256 of the bytes."""
    target = Path(path)
    hasher = hashlib.sha256()
    with target.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            hasher.update(chunk)
    stat = target.stat()
    stamp = f"{stat.st_size}:{int(stat.st_mtime)}:{hasher.hexdigest()}"
    return "sha256:" + hashlib.sha256(stamp.encode("utf-8")).hexdigest()


def fingerprint_region(ws, ref: str) -> str:
    """Deterministic hash of cell values + formula texts in one range.

    Empty cells are normalised so that a trailing blank never changes the
    digest. This is the identity used for idempotency decisions (I4).
    """
    from openpyxl.utils import range_boundaries  # noqa: PLC0415

    min_col, min_row, max_col, max_row = range_boundaries(ref)
    parts = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                            min_col=min_col, max_col=max_col):
        for cell in row:
            value = cell.value
            if value is None:
                parts.append(f"{cell.coordinate}=")
            else:
                parts.append(f"{cell.coordinate}={type(value).__name__}:{value}")
    joined = "\x1f".join(parts)
    return "sha256:" + hashlib.sha256(joined.encode("utf-8")).hexdigest()


def fingerprint_structure_parts(sheet_parts) -> str:
    """Structural identity from canonical 'name|dims|tables' parts."""
    joined = "\x1f".join(sorted(sheet_parts))
    return "sha256:" + hashlib.sha256(joined.encode("utf-8")).hexdigest()


def fingerprint_structure(wb) -> str:
    """Sheet names + table refs + used ranges (structural identity).

    Note: openpyxl 3.1's TableList.items() yields (name, ref-string) pairs
    -- the second item is already the ref, not a Table object.
    """
    parts = []
    for ws in wb.worksheets:
        tables = ",".join(
            f"{name}={ref}"
            for name, ref in sorted(getattr(ws, "tables", {}).items())
        )
        parts.append(f"{ws.title}|{ws.dimensions}|{tables}")
    return fingerprint_structure_parts(parts)


# --------------------------------------------------------------------------
# backup / temp / atomic commit  (I5)
# --------------------------------------------------------------------------

def backup_workbook(path, *, tag: str = "") -> dict:
    """Copy a workbook aside before it is modified. Fail-closed."""
    source = Path(path)
    if not source.exists():
        raise XlsxError("BACKUP_FAILED", f"Cannot back up missing file: {source}",
                        context={"path": str(source)})
    fault("backup", "backup requested before a write", path=str(source))
    stamp = _timestamp()
    suffix = f".bak-{stamp}" + (f"-{tag}" if tag else "") + source.suffix
    target = source.with_name(source.stem + suffix)
    try:
        shutil.copy2(source, target)
    except OSError as exc:
        raise XlsxError(
            "BACKUP_FAILED",
            f"Could not create backup next to {source}: {exc}",
            recovery="Check free disk space and folder write permissions. "
                     "Nothing was written.",
            context={"path": str(source), "target": str(target)},
        ) from exc
    return {"ok": True, "backup": str(target),
            "sha256": fingerprint_file(target), "bytes": target.stat().st_size}


def _timestamp() -> str:
    from datetime import datetime  # noqa: PLC0415
    return datetime.now().strftime("%Y%m%dT%H%M%S")


def _new_operation_id(prefix: str = "op") -> str:
    import secrets  # noqa: PLC0415
    return f"{prefix}_{_timestamp()}_{secrets.token_hex(2)}"


def _same_filesystem(left: Path, right: Path) -> bool:
    """True when both paths live on one device (os.replace can be atomic)."""
    try:
        return left.stat().st_dev == right.stat().st_dev
    except OSError:
        return True          # unknown -> attempt the atomic path first


def atomic_replace(tmp_path, target) -> dict:
    """Replace `target` with `tmp_path` atomically where the OS allows it.

    Failure is fail-closed and fully diagnosed: the destination is left
    untouched, the staged file (when it still exists) is kept for
    inspection, and the error carries ``temp_present`` / ``committed`` /
    ``commit_state`` so a caller never has to guess what happened.
    """
    tmp, dest = Path(tmp_path), Path(target)
    fault("atomic_replace", "about to replace the destination",
          tmp=str(tmp), target=str(dest))
    if not tmp.exists():
        raise XlsxError(
            "ATOMIC_COMMIT_FAILED",
            f"the staged temporary file is gone: {tmp}",
            recovery=("Nothing was committed. Re-run the execution to stage "
                      "a new output file."),
            context={"tmp": str(tmp), "target": str(dest), "temp_present": False,
                     "committed": False, "commit_state": "FAILED"})
    same_fs = _same_filesystem(tmp.parent, dest.parent)
    try:
        os.replace(tmp, dest)
        return {"ok": True, "atomic": True, "output": str(dest),
                "same_filesystem": same_fs, "temp_present": False,
                "committed": True}
    except OSError as replace_exc:
        if same_fs:
            raise XlsxError(
                "ATOMIC_COMMIT_FAILED",
                f"could not place the staged output over the destination: "
                f"{replace_exc}",
                recovery=("The original file was never modified and the "
                          "staged file was kept for diagnostics. Close the "
                          "program holding the file and re-run."),
                context={"tmp": str(tmp), "target": str(dest),
                         "temp_present": True, "committed": False,
                         "commit_state": "FAILED", "same_filesystem": True,
                         "diagnostics": [f"{type(replace_exc).__name__}: "
                                         f"{replace_exc}"]}) from replace_exc
        # Cross-device fallback: move, then let the caller verify by re-opening.
        try:
            shutil.move(str(tmp), str(dest))
        except OSError as exc:
            raise XlsxError(
                "ATOMIC_COMMIT_FAILED",
                f"could not place the output after the cross-device move: {exc}",
                recovery=("The original file was never modified and the "
                          "staged file was kept for diagnostics."),
                context={"tmp": str(tmp), "target": str(dest),
                         "temp_present": Path(tmp).exists(), "committed": False,
                         "commit_state": "FAILED", "same_filesystem": False,
                         "diagnostics": [f"{type(exc).__name__}: {exc}"]}) from exc
        return {"ok": True, "atomic": False, "output": str(dest),
                "same_filesystem": False, "temp_present": True,
                "committed": True,
                "diagnostics": ["cross-device move used instead of os.replace"]}


def verify_workbook(path, *, expect_sheets=None) -> dict:
    """Re-open a written workbook and check it is structurally sound."""
    from openpyxl import load_workbook  # noqa: PLC0415

    target = Path(path)
    if not target.exists():
        return {"ok": False, "reason": "output missing"}
    try:
        wb = load_workbook(str(target), read_only=True)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False,
                "reason": f"{type(exc).__name__}: {exc}"}
    try:
        sheets = list(wb.sheetnames)
        if not sheets:
            return {"ok": False, "reason": "workbook has no sheets"}
        if expect_sheets is not None and sorted(sheets) != sorted(expect_sheets):
            return {"ok": False,
                    "reason": "sheet list changed",
                    "expected": sorted(expect_sheets), "actual": sorted(sheets)}
        return {"ok": True, "sheets": sheets, "sheet_count": len(sheets)}
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass


def _plan_hash(plan) -> str:
    """Deterministic hash of a caller-supplied operation plan (I4 seed)."""
    canonical = json.dumps(plan, sort_keys=True, ensure_ascii=False,
                           default=_json_default)
    return "sha256:" + hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def save_workbook_safe(loaded: LoadedWorkbook, out_path=None, *,
                       backup: bool = False, in_place: bool = False,
                       approve_token: str = None,
                       require_approval: bool = False,
                       dry_run: bool = False,
                       expect_sheets=None,
                       plan=None, region_fingerprint=None,
                       manifest_dir=None, manifest_status: str = "committed",
                       record_manifest: bool = True) -> dict:
    """The single write path: backup -> temp -> validate -> atomic replace.

    Gates, all fail-closed (the original file never changes on any failure):

      1. values_only / read_only handles are refused outright (I1 + I2).
      2. An unacknowledged in-place write (target == existing source without
         ``in_place=True``) is refused -- in-place is never the default (I5).
      3. When the caller declares ``require_approval=True`` (plan-based flows),
         an in-place write additionally requires a non-empty ``approve_token``
         (Phase-3 contract: first 12 chars of the plan hash). Legacy CLI flows
         carry approval at the HERMES call site and pass
         ``require_approval=False``; their in-place write is recorded as
         ``caller-explicit`` in the result and manifest.
      4. Every existing file the write could destroy is backed up first;
         any backup failure aborts with BACKUP_FAILED.
      5. The temp output must re-open cleanly before it replaces anything;
         failure keeps the temp file for diagnostics and aborts.

    After a successful commit the operation is recorded in the skill's
    manifest (.xlsx_ops/manifest.jsonl) with operation_id, plan_hash (when
    the caller passes ``plan``), source/target fingerprints, optional target
    region fingerprint, timestamp, status and output. Manifest problems are
    reported in the returned dict, never silenced.
    """
    # --- gate 1: the handle must carry write authority (I1 + I2) ----------
    if loaded.mode == "values_only":
        raise XlsxError(
            "DATA_ONLY_WRITE_FORBIDDEN",
            "A workbook opened with data_only=True cannot be saved: every "
            "formula would be replaced by its last cached value and lost "
            "permanently.",
            recovery="Re-open the file with data_only=False for the write pass; "
                     "keep this handle for reading values only.",
            context={"path": loaded.path, "mode": loaded.mode},
        )
    if loaded.mode == "read_only":
        raise XlsxError(
            "READ_ONLY_FILE",
            "A workbook opened with read_only=True cannot be saved.",
            recovery="Load the file a second time with read_only=False (Pass C) "
                     "and apply the same plan to that handle.",
            context={"path": loaded.path, "mode": loaded.mode},
        )

    source = Path(loaded.path).resolve() if loaded.path else None
    if out_path is None:
        if source is None:
            raise XlsxError(
                "WRITE_FAILED",
                "No output path given and the handle carries no source path.",
                context={"mode": loaded.mode})
        # Default policy (I5): never overwrite the source -- derive a new name.
        target = source.with_name(f"{source.stem}_filled{source.suffix}")
    else:
        target = Path(out_path).resolve()

    writing_in_place = bool(source and target == source and source.exists())
    if writing_in_place and not in_place:
        raise XlsxError(
            "READ_ONLY_FILE",
            "Refusing to overwrite the source file in place without an "
            "explicit in-place acknowledgement.",
            recovery=("Pass in_place=True / --in-place to confirm the target, "
                      "or write to a new file (the default)."),
            context={"source": str(source)},
        )
    if writing_in_place and require_approval and not approve_token:
        raise XlsxError(
            "PLAN_CONFLICT",
            "This operation requires explicit approval before it may "
            "overwrite the source in place.",
            recovery=("Re-run with --approve-token <first 12 chars of the "
                      "plan hash> after the user confirms the plan."),
            context={"source": str(source)},
        )
    if approve_token and not writing_in_place:
        # Token present but nothing approved: keep it honest in the result.
        pass

    approval = "not-required"
    if writing_in_place:
        approval = ("token:" + approve_token if approve_token
                    else "caller-explicit")
    plan = {
        "source": str(source) if source else None,
        "target": str(target),
        "in_place": writing_in_place,
        "backup": bool(backup or writing_in_place),
        "approval": approval,
    }

    if dry_run:
        return {"ok": True, "dry_run": True, "would_write": str(target),
                "in_place": writing_in_place, "bytes_written": 0,
                "approval": approval, "plan": plan}

    # --- gate 2: back up everything the write could destroy ---------------
    operation_id = _new_operation_id()
    source_before = fingerprint_file(source) if source and source.exists() else None
    backups = []
    if plan["backup"]:
        candidates = []
        if source and source.exists():
            candidates.append(source)
        if target.exists() and target != source:
            candidates.append(target)
        for candidate in candidates:
            info = backup_workbook(candidate, tag="prewrite")
            backups.append(info)

    # --- gate 3: write to a temp file in the target directory -------------
    fault("staged_save", "before writing the temporary output",
          target=str(target))
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp_handle, tmp_name = tempfile.mkstemp(
        prefix=f".{target.stem}.tmp-{operation_id}-", suffix=".xlsx",
        dir=str(target.parent))
    os.close(tmp_handle)
    try:
        loaded.wb.save(tmp_name)
    except Exception as exc:  # noqa: BLE001
        raise XlsxError(
            "WRITE_FAILED",
            f"Writing the temporary output failed: {type(exc).__name__}: {exc}",
            recovery=("The original file was not modified"
                      + (f" and a backup exists at {backups[0]['backup']}"
                         if backups else "") + ". The temporary file was kept "
                      "for diagnostics."),
            context={"tmp": tmp_name,
                     "backups": [b["backup"] for b in backups]},
        ) from exc

    # --- gate 4: the temp output must reopen cleanly ----------------------
    fault("pre_commit_validate", "before re-opening the temporary output",
          tmp=tmp_name)
    probe = verify_workbook(tmp_name, expect_sheets=expect_sheets)
    if not probe.get("ok"):
        raise XlsxError(
            "VALIDATION_FAILED",
            f"Temporary output did not validate: {probe.get('reason')}",
            recovery=("Nothing was committed. The temporary file was kept for "
                      "inspection."),
            context={"tmp": tmp_name, "probe": probe,
                     "backups": [b["backup"] for b in backups]},
        )

    # --- gate 5: commit + verify ------------------------------------------
    placed = atomic_replace(tmp_name, target)
    final = verify_workbook(target, expect_sheets=expect_sheets)
    if not final.get("ok"):
        raise XlsxError("WRITE_FAILED",
                        f"Committed file failed re-verification: "
                        f"{final.get('reason')}",
                        context={"output": str(target),
                                 "backups": [b["backup"] for b in backups]})

    # Provenance (I4 foundation): record the committed operation. Manifest
    # problems surface in the returned dict, never silenced.
    target_after = fingerprint_file(target)
    manifest_entry = {
        "operation_id": operation_id,
        "plan_hash": _plan_hash(plan) if plan is not None else None,
        "source_fingerprint": source_before,
        "target_fingerprint": target_after,
        "target_region_fingerprint": region_fingerprint,
        "region_fingerprint": region_fingerprint,
        "timestamp": _timestamp(),
        # FAZ 3C (D1): only a real commit may claim "committed"; a staged
        # intermediate output records "staged" and is never mistaken for a
        # completed operation by the idempotency layer.
        "status": manifest_status,
        "output": str(target),
        "in_place": writing_in_place,
        "backup": backups[0]["backup"] if backups else None,
    }
    if record_manifest:
        manifest_info = manifest_append(manifest_entry, skill_dir=manifest_dir)
    else:
        manifest_info = {"ok": True, "skipped": True,
                         "reason": "record_manifest=False (staged intermediate)"}

    return {
        "ok": True,
        "output": str(target),
        "bytes": target.stat().st_size,
        "atomic": placed["atomic"],
        "bytes_written": target.stat().st_size,
        "operation_id": operation_id,
        "plan_hash": manifest_entry["plan_hash"],
        "in_place": writing_in_place,
        "approval": approval,
        "backup": backups[0]["backup"] if backups else None,
        "backups": [b["backup"] for b in backups],
        "backup_fingerprint": backups[0]["sha256"] if backups else None,
        "source_fingerprint_before": source_before,
        "target_fingerprint_after": target_after,
        "sheets": final["sheets"],
        "verified": True,
        "manifest": manifest_info,
        "manifest_status": manifest_entry["status"],
    }


def _try_unlink(path) -> None:
    try:
        os.unlink(path)
    except OSError:
        pass


# --------------------------------------------------------------------------
# operation manifest  (idempotency foundation)
# --------------------------------------------------------------------------

MANIFEST_LIMIT = 200


def manifest_path(skill_dir=None) -> Path:
    base = Path(skill_dir) if skill_dir else Path(__file__).resolve().parent.parent
    return base / ".xlsx_ops" / "manifest.jsonl"


class _ManifestLock:
    """Exclusive lock around one manifest read-modify-write cycle (D5).

    A sibling ``.lock`` file is opened and locked with an OS byte-range
    lock (``msvcrt.locking`` on Windows, ``fcntl.flock`` elsewhere) so two
    concurrent executions can never interleave a half-written line or lose
    an entry to the trim rewrite. Minimal by design: no distributed DB,
    no daemon -- one lock file next to the manifest.
    """

    def __init__(self, path: Path):
        self.path = Path(str(path) + ".lock")
        self.fd = None
        self.acquired = False

    def __enter__(self):
        for attempt in range(21):
            try:
                self.fd = os.open(str(self.path), os.O_CREAT | os.O_RDWR)
                if os.name == "nt":
                    import msvcrt  # noqa: PLC0415
                    msvcrt.locking(self.fd, msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl  # noqa: PLC0415
                    fcntl.flock(self.fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                self.acquired = True
                return self
            except OSError:
                if self.fd is not None:
                    try:
                        os.close(self.fd)
                    except OSError:
                        pass
                    self.fd = None
                if attempt < 20:
                    time.sleep(0.05)
        return self

    def __exit__(self, *exc_info):
        if self.fd is not None:
            try:
                if os.name == "nt":
                    import msvcrt  # noqa: PLC0415
                    msvcrt.locking(self.fd, msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl  # noqa: PLC0415
                    fcntl.flock(self.fd, fcntl.LOCK_UN)
            except OSError:
                pass
            try:
                os.close(self.fd)
            except OSError:
                pass
            self.fd = None
        return False


def manifest_append(entry: dict, *, skill_dir=None) -> dict:
    """Append one operation record; keep the newest MANIFEST_LIMIT lines.

    The whole append+trim cycle runs under ``_ManifestLock`` (D5). A lock
    that cannot be taken within the retry budget is reported as
    ``MANIFEST_LOCKED`` instead of writing a possibly-interleaved line.
    """
    path = manifest_path(skill_dir)
    fault("manifest_append", "manifest append requested", path=str(path))
    record = dict(entry)
    record.setdefault("timestamp", _timestamp())
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return {"ok": False, "error_code": "MANIFEST_WRITE_FAILED",
                "message": f"Manifest directory could not be created: {exc}",
                "context": {"path": str(path)}}
    with _ManifestLock(path) as lock:
        if not lock.acquired:
            return {"ok": False, "error_code": "MANIFEST_LOCKED",
                    "message": ("Another execution holds the manifest lock; "
                                "nothing was recorded."),
                    "context": {"path": str(path), "lock": str(lock.path)},
                    "diagnostics": ["MANIFEST_LOCKED: append skipped"]}
        try:
            with path.open("a", encoding="utf-8") as fh:
                fh.write(json.dumps(record, ensure_ascii=False,
                                    default=_json_default) + "\n")
        except OSError as exc:
            return {"ok": False, "error_code": "MANIFEST_WRITE_FAILED",
                    "message": f"Manifest append failed: {exc}",
                    "context": {"path": str(path)}}
        trim = _trim_manifest(path)
    result = {"ok": True, "path": str(path),
              "operation_id": record.get("operation_id"),
              "status_recorded": record.get("status"), "trim": trim}
    if trim.get("trimmed"):
        result["diagnostics"] = [f"MANIFEST_TRIM_REPORTED: kept the newest "
                                 f"{MANIFEST_LIMIT} lines "
                                 f"({trim.get('removed')} removed)"]
    return result


def _trim_manifest(path: Path) -> dict:
    """Trim to MANIFEST_LIMIT lines, reporting what happened (never silent)."""
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        return {"trimmed": False, "removed": 0,
                "diagnostics": [f"manifest trim skipped: {exc}"]}
    if len(lines) <= MANIFEST_LIMIT:
        return {"trimmed": False, "removed": 0,
                "diagnostics": ["MANIFEST_TRIM_REPORTED: nothing to trim"],
                "lines": len(lines)}
    kept = lines[-MANIFEST_LIMIT:]
    dropped = lines[:len(lines) - MANIFEST_LIMIT]
    dropped_corrupt = sum(1 for line in dropped if not _line_is_json(line))
    try:
        path.write_text("\n".join(kept) + "\n", encoding="utf-8")
    except OSError as exc:
        return {"trimmed": False, "removed": 0,
                "diagnostics": [f"manifest trim failed: {exc}"]}
    return {"trimmed": True, "removed": len(dropped),
            "dropped_corrupt_lines": dropped_corrupt,
            "lines": len(kept)}


def _line_is_json(line: str) -> bool:
    try:
        json.loads(line)
        return True
    except json.JSONDecodeError:
        return False


def manifest_read(*, skill_dir=None) -> dict:
    """Read the manifest, reporting corruption instead of failing open."""
    path = manifest_path(skill_dir)
    if not path.exists():
        dir_exists = path.parent.exists()
        return {"ok": True, "entries": [], "status": "missing",
                "error_code": "MANIFEST_MISSING", "corrupt_lines": [],
                "dir_exists": dir_exists,
                "diagnostics": [
                    {"code": "MANIFEST_MISSING",
                     "message": ("manifest does not exist yet; no idempotency "
                                 "evidence is available"
                                 + ("" if dir_exists else
                                    " (fresh workspace: the manifest was never "
                                    "created here)")),
                     "context": {"path": str(path),
                                 "dir_exists": dir_exists}}]}
    entries, broken = [], []
    try:
        raw_lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        return {"ok": False, "error_code": "MANIFEST_MISSING",
                "status": "unreadable", "corrupt_lines": [],
                "message": f"Manifest unreadable: {exc}",
                "diagnostics": [
                    {"code": "MANIFEST_MISSING",
                     "message": ("manifest could not be read; idempotency "
                                 "evidence is unavailable"),
                     "context": {"path": str(path), "error": str(exc)}}],
                "context": {"path": str(path)}}
    for index, line in enumerate(raw_lines, start=1):
        if not line.strip():
            continue
        try:
            entries.append(json.loads(line))
        except json.JSONDecodeError:
            broken.append(index)
    diagnostics = []
    if broken:
        diagnostics.append({
            "code": "MANIFEST_CORRUPT",
            "message": f"{len(broken)} manifest line(s) are not valid JSON and "
                       "were skipped",
            "context": {"lines": broken[:10], "path": str(path)},
        })
    return {"ok": True, "entries": entries, "diagnostics": diagnostics,
            "corrupt_lines": broken, "corrupt_count": len(broken),
            "status": "corrupt" if broken else "ok"}


def _is_committed(entry: dict) -> bool:
    """Only a real commit is idempotency evidence (FAZ 3C, D1).

    Staged intermediates (``status="staged"``) and failed runs
    (``status="failed"``) are diagnostics, never proof that the plan was
    applied. Records written before the status field existed stay eligible.
    """
    if "status" not in entry:
        return True
    return entry.get("status") == "committed"


def manifest_matches(plan_hash: str, region_fingerprint: str = None, *,
                     read: dict = None, skill_dir=None) -> list:
    """Every committed record for plan_hash (optionally one region)."""
    read = read if read is not None else manifest_read(skill_dir=skill_dir)
    hits = [entry for entry in read.get("entries", [])
            if _is_committed(entry) and entry.get("plan_hash") == plan_hash]
    if region_fingerprint:
        hits = [entry for entry in hits
                if entry.get("region_fingerprint") == region_fingerprint]
    return hits


def manifest_find(plan_hash: str, region_fingerprint: str = None, *,
                  skill_dir=None, read: dict = None) -> dict | None:
    """Newest committed record matching plan_hash (and region when given)."""
    hits = manifest_matches(plan_hash, region_fingerprint, read=read,
                            skill_dir=skill_dir)
    return hits[-1] if hits else None


def canonical_json(obj) -> str:
    """Deterministic JSON text: sorted keys, UTF-8 kept, no spaces."""
    return json.dumps(obj, ensure_ascii=False, sort_keys=True,
                      separators=(",", ":"))


def sha256_text(text: str) -> str:
    """sha256 of a UTF-8 string (identity helper for FAZ 2B)."""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def sha256_file(path) -> dict:
    """Pure content identity of a file: sha256 + size, NO mtime.

    Deliberately differs from fingerprint_file(): identity that must
    survive mtime changes (FAZ 2B profiles) cannot hash the mtime.
    """
    target = Path(path)
    hasher = hashlib.sha256()
    with target.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            hasher.update(chunk)
    return {"sha256": hasher.hexdigest(), "size": target.stat().st_size}


def idempotency_decision(plan_hash: str, region_fp: str, *,
                         current_region_empty: bool = False,
                         skill_dir=None) -> dict:
    """I4 decision: already_applied / safe_to_reapply / conflict /
    first_run / unknown (FAZ 3C).

    ``unknown`` means the manifest could not prove anything (it is missing,
    unreadable, or carries unparsable lines). Per decision D3 the caller
    must fail closed on a non-empty target region instead of assuming a
    first run -- absent evidence is never treated as evidence of absence.
    """
    read = manifest_read(skill_dir=skill_dir)
    manifest_status = read.get("status", "ok")
    corrupt = list(read.get("corrupt_lines") or [])
    exact = manifest_matches(plan_hash, region_fp, read=read)
    if exact:
        decision = {"status": "already_applied", "prior": exact[-1],
                    "match_count": len(exact),
                    "message": "This exact plan was already applied to this "
                               "region."}
        if len(exact) > 1:
            decision["duplicate_entry"] = True
            decision["warning_code"] = "MANIFEST_DUPLICATE"
            decision["message"] += (f" ({len(exact)} identical records; the "
                                    "newest one was used)")
        return decision
    any_hits = manifest_matches(plan_hash, None, read=read)
    if any_hits:
        any_plan = any_hits[-1]
        prior_before = any_plan.get("region_fingerprint_before")
        in_original_state = bool(region_fp and prior_before
                                 and region_fp == prior_before)
        if current_region_empty or in_original_state:
            return {"status": "safe_to_reapply", "prior": any_plan,
                    "match_count": len(any_hits),
                    "message": "The plan ran before and the target region is "
                               "in its pre-execution state (empty or "
                               "unchanged); re-applying is safe."}
        return {"status": "conflict", "prior": any_plan,
                    "match_count": len(any_hits),
                    "message": "The plan ran before and the target region holds "
                               "different content; stopping for review.",
                    "error_code": "PLAN_CONFLICT"}
    if manifest_status in ("missing", "unreadable") or corrupt:
        code = "MANIFEST_CORRUPT" if corrupt else "MANIFEST_MISSING"
        fresh = (manifest_status == "missing"
                 and not read.get("dir_exists", True))
        return {"status": "unknown", "prior": None, "error_code": code,
                "fresh_workspace": fresh,
                "evidence": {"manifest_status": manifest_status,
                             "corrupt_lines": corrupt,
                             "fresh_workspace": fresh,
                             "entries_read": len(read.get("entries") or [])},
                "message": ("No idempotency evidence is available for this "
                            "plan (manifest " + manifest_status +
                            "); a first run cannot be distinguished from a "
                            "re-run, so the caller must fail closed unless "
                            "the target region is untouched.")} 
    return {"status": "first_run", "prior": None,
            "message": "No prior execution recorded for this plan."}
