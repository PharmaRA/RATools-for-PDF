import argparse
import json
import os
import shlex
import sys
import traceback
import xml.etree.ElementTree as ET
from datetime import datetime


KEYWORDS = (
    "preflight",
    "fix",
    "save",
    "menu",
    "js",
    "script",
    "action",
    "batch",
    "wizard",
    "print",
    "exec",
)

MENU_CANDIDATES = (
    "PrintProduction",
    "Preflight",
    "Preflight:Preflight",
    "DocumentPreflight",
    "DocumentProperties",
)

JS_CANDIDATES = (
    "app",
    "execMenuItem",
    "beginPriv",
    "endPriv",
    "trustedFunction",
    "preflight",
)

TRANSITION_MENU_CANDIDATES = (
    "PrintProduction",
    "Preflight",
)

WORKFLOW_SEQUENCE_PATH = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Sequences\ENU\Action04.sequ"
WORKFLOW_REGISTRY_PATH = r"HKCU\Software\Adobe\Adobe Acrobat\DC\Workflow\cRegistered\c3"
WORKFLOW_REGISTRY_ROOT = r"Software\Adobe\Adobe Acrobat\DC\Workflow\cRegistered"
SEQUENCE_FILE_ASSOC_ROOT = r"Acrobat.Sequence\shell\Import_Action\command"


def list_public_member_names(obj):
    names = []
    for name in dir(obj):
        if name.startswith("_"):
            continue
        names.append(name)
    return sorted(set(names), key=str.lower)


def find_interesting_members(member_names):
    interesting = []
    for name in member_names:
        lowered = name.lower()
        if any(keyword in lowered for keyword in KEYWORDS):
            interesting.append(name)
    return sorted(interesting, key=str.lower)


def _safe_call(callable_obj):
    try:
        value = callable_obj()
        if isinstance(value, (str, int, float, bool)) or value is None:
            serializable_value = value
        else:
            serializable_value = {
                "kind": type(value).__name__,
                "repr": repr(value),
            }
        return {"ok": True, "value": serializable_value, "raw": value}
    except Exception as exc:
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}


def extract_sequence_import_command(command_template, sequence_path):
    expanded = command_template.replace("%1", sequence_path)
    return [part.strip('"') for part in shlex.split(expanded, posix=False)]


def extract_sequence_metadata(xml_text):
    namespace = {"wf": "http://ns.adobe.com/acrobat/workflow/2012"}
    root = ET.fromstring(xml_text)
    metadata = {
        "title": root.attrib.get("title", ""),
        "description": root.attrib.get("description", ""),
        "preflight_profiles": [],
    }

    for command in root.findall(".//wf:Command[@name='CALS:Preflight']", namespace):
        items = {}
        for item in command.findall("wf:Items/wf:Item", namespace):
            item_name = item.attrib.get("name", "")
            items[item_name] = item.attrib.get("value", "")
        metadata["preflight_profiles"].append(
            {
                "name": items.get("CALS_PREFLIGHT_CMD_PROFILE_NAME", ""),
                "dictkey": items.get("CALS_PREFLIGHT_CMD_PROFILE_DICTKEY", ""),
                "fingerprint": items.get("CALS_PREFLIGHT_CMD_PROFILE_FINGERPRINT", ""),
            }
        )

    return metadata


def build_workflow_probe_candidates(metadata, workflow_relative_path="", workflow_id=""):
    candidates = []

    def add_candidate(value):
        if value and value not in candidates:
            candidates.append(value)

    add_candidate(metadata.get("title", ""))
    add_candidate(workflow_relative_path)
    add_candidate(os.path.splitext(workflow_relative_path)[0] if workflow_relative_path else "")
    add_candidate(workflow_id)

    for profile in metadata.get("preflight_profiles", []):
        add_candidate(profile.get("name", ""))
        add_candidate(profile.get("dictkey", ""))
        add_candidate(profile.get("fingerprint", ""))

    return candidates


def build_workflow_registry_index(entries):
    index = []
    for entry in entries:
        values = entry.get("values", {})
        index.append(
            {
                "key_name": entry.get("key_name", ""),
                "title": values.get("tTitle", ""),
                "id": values.get("aID", ""),
                "relative_path": values.get("tRelativeDIPath", ""),
            }
        )
    return index


def _read_registry_default_value(root, path):
    try:
        import winreg
        hive = getattr(winreg, root)
        with winreg.OpenKey(hive, path) as key:
            value, _value_type = winreg.QueryValueEx(key, "")
            return {"ok": True, "value": value}
    except Exception as exc:
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}


def _enumerate_workflow_registry_entries():
    try:
        import winreg
        entries = []
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, WORKFLOW_REGISTRY_ROOT) as root_key:
            index = 0
            while True:
                try:
                    child_name = winreg.EnumKey(root_key, index)
                except OSError:
                    break
                with winreg.OpenKey(root_key, child_name) as child_key:
                    values = {}
                    value_index = 0
                    while True:
                        try:
                            value_name, value_data, _value_type = winreg.EnumValue(child_key, value_index)
                        except OSError:
                            break
                        values[value_name] = value_data
                        value_index += 1
                entries.append({"key_name": child_name, "values": values})
                index += 1
        return {"ok": True, "entries": entries}
    except Exception as exc:
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}


def probe_menu_candidates(app, candidate_names, execute_enabled=False):
    result = {}
    for name in candidate_names:
        enabled = _safe_call(lambda current_name=name: getattr(app, "MenuItemIsEnabled")(current_name))
        entry = {"enabled": {key: value for key, value in enabled.items() if key != "raw"}}
        if execute_enabled and enabled.get("ok") and enabled.get("raw"):
            executed = _safe_call(lambda current_name=name: getattr(app, "MenuItemExecute")(current_name))
            entry["execute"] = {key: value for key, value in executed.items() if key != "raw"}
        result[name] = entry
    return result


def probe_js_candidates(js_obj, candidate_names):
    result = {}
    for name in candidate_names:
        try:
            value = getattr(js_obj, name)
        except AttributeError:
            result[name] = {"exists": False}
            continue
        except Exception as exc:
            result[name] = {
                "exists": False,
                "access_error": f"{type(exc).__name__}: {exc}",
            }
            continue

        entry = {
            "exists": True,
            "callable": callable(value),
            "value": {key: value_ for key, value_ in _safe_call(lambda current=value: current).items() if key != "raw"},
        }
        if callable(value):
            if name == "execMenuItem":
                call_test = _safe_call(lambda current=value: current("PrintProduction"))
            else:
                call_test = _safe_call(lambda current=value: current())
            entry["call_test"] = {key: value_ for key, value_ in call_test.items() if key != "raw"}
        result[name] = entry
    return result


def probe_app_state(app):
    return {
        "GetActiveTool": {key: value for key, value in _safe_call(lambda: getattr(app, "GetActiveTool")()).items() if key != "raw"},
        "GetNumAVDocs": {key: value for key, value in _safe_call(lambda: getattr(app, "GetNumAVDocs")()).items() if key != "raw"},
        "GetActiveDoc": {key: value for key, value in _safe_call(lambda: getattr(app, "GetActiveDoc")()).items() if key != "raw"},
    }


def probe_menu_transition(app, menu_name):
    return {
        "before": probe_app_state(app),
        "execute": {key: value for key, value in _safe_call(lambda: getattr(app, "MenuItemExecute")(menu_name)).items() if key != "raw"},
        "after": probe_app_state(app),
    }


def _read_workflow_registry_value(value_name):
    result = _safe_call(
        lambda: __import__("winreg").OpenKey(
            __import__("winreg").HKEY_CURRENT_USER,
            WORKFLOW_REGISTRY_PATH.replace("HKCU\\", ""),
        )
    )
    if not result.get("ok") or result.get("raw") is None:
        return {key: value for key, value in result.items() if key != "raw"}

    winreg = __import__("winreg")
    key = result["raw"]
    try:
        value, _value_type = winreg.QueryValueEx(key, value_name)
        return {"ok": True, "value": value}
    except Exception as exc:
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    finally:
        winreg.CloseKey(key)


def _describe_com_object(obj):
    if obj is None:
        return {"kind": "null"}
    try:
        members = list_public_member_names(obj)
    except Exception as exc:
        return {
            "kind": type(obj).__name__,
            "member_count": None,
            "interesting_members": [],
            "sample_members": [],
            "introspection_error": f"{type(exc).__name__}: {exc}",
        }
    return {
        "kind": type(obj).__name__,
        "member_count": len(members),
        "interesting_members": find_interesting_members(members),
        "sample_members": members[:80],
    }


def probe_acrobat(pdf_path=None):
    try:
        import win32com.client
        import pythoncom
    except Exception as exc:
        return {
            "ok": False,
            "stage": "import",
            "error": f"{type(exc).__name__}: {exc}",
        }

    pythoncom.CoInitialize()
    app = None
    av_doc = None
    pd_doc = None
    try:
        app = win32com.client.Dispatch("AcroExch.App")
        app_members = list_public_member_names(app)

        result = {
            "ok": True,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "pdf_path": pdf_path or "",
            "workflow_probe": {},
            "objects": {
                "AcroExch.App": {
                    "member_count": len(app_members),
                    "interesting_members": find_interesting_members(app_members),
                    "sample_members": app_members[:80],
                    "property_checks": {
                        "Show": _safe_call(lambda: getattr(app, "Show")()),
                        "Hide": _safe_call(lambda: getattr(app, "Hide")()),
                        "GetNumAVDocs": _safe_call(lambda: getattr(app, "GetNumAVDocs")()),
                    },
                }
            },
        }

        if os.path.exists(WORKFLOW_SEQUENCE_PATH):
            with open(WORKFLOW_SEQUENCE_PATH, "r", encoding="utf-8") as seq_file:
                sequence_metadata = extract_sequence_metadata(seq_file.read())
            workflow_relative_path = _read_workflow_registry_value("tRelativeDIPath").get("value", "")
            workflow_id = _read_workflow_registry_value("aID").get("value", "")
            sequence_import_command = _read_registry_default_value("HKEY_CLASSES_ROOT", SEQUENCE_FILE_ASSOC_ROOT)
            workflow_registry_entries = _enumerate_workflow_registry_entries()
            workflow_candidates = build_workflow_probe_candidates(
                sequence_metadata,
                workflow_relative_path=workflow_relative_path,
                workflow_id=workflow_id,
            )
            result["workflow_probe"] = {
                "sequence_path": WORKFLOW_SEQUENCE_PATH,
                "registry_key": WORKFLOW_REGISTRY_PATH,
                "sequence_metadata": sequence_metadata,
                "registry_values": {
                    "tRelativeDIPath": _read_workflow_registry_value("tRelativeDIPath"),
                    "aID": _read_workflow_registry_value("aID"),
                    "tTitle": _read_workflow_registry_value("tTitle"),
                },
                "sequence_import": {
                    "registry_command": sequence_import_command,
                    "expanded_command": extract_sequence_import_command(
                        sequence_import_command.get("value", ""),
                        WORKFLOW_SEQUENCE_PATH,
                    ) if sequence_import_command.get("ok") else [],
                },
                "registered_workflows": build_workflow_registry_index(workflow_registry_entries.get("entries", []))
                if workflow_registry_entries.get("ok")
                else [],
                "probe_candidates": workflow_candidates,
            }

        av_doc = win32com.client.Dispatch("AcroExch.AVDoc")
        av_members = list_public_member_names(av_doc)
        result["objects"]["AcroExch.AVDoc"] = {
            "member_count": len(av_members),
            "interesting_members": find_interesting_members(av_members),
            "sample_members": av_members[:80],
            "property_checks": {
                "Open": _safe_call(lambda: getattr(av_doc, "Open")("", "")),
                "GetPDDoc": _safe_call(lambda: getattr(av_doc, "GetPDDoc")()),
                "BringToFront": _safe_call(lambda: getattr(av_doc, "BringToFront")()),
            },
        }

        pd_doc = win32com.client.Dispatch("AcroExch.PDDoc")
        pd_members = list_public_member_names(pd_doc)
        result["objects"]["AcroExch.PDDoc"] = {
            "member_count": len(pd_members),
            "interesting_members": find_interesting_members(pd_members),
            "sample_members": pd_members[:80],
            "property_checks": {
                "GetNumPages": _safe_call(lambda: getattr(pd_doc, "GetNumPages")()),
                "GetJSObject": _safe_call(lambda: getattr(pd_doc, "GetJSObject")()),
                "Save": _safe_call(lambda: getattr(pd_doc, "Save")(1, "")),
            },
        }

        if pdf_path:
            normalized_pdf_path = os.path.abspath(pdf_path)
            open_result = _safe_call(lambda: getattr(av_doc, "Open")(normalized_pdf_path, "Acrobat Probe"))
            result["document_probe"] = {
                "open_result": open_result,
                "input_exists": os.path.exists(normalized_pdf_path),
                "input_path": normalized_pdf_path,
            }
            if open_result.get("ok") and open_result.get("value"):
                active_pd_doc = _safe_call(lambda: getattr(av_doc, "GetPDDoc")())
                result["document_probe"]["GetPDDoc"] = {
                    key: value for key, value in active_pd_doc.items() if key != "raw"
                }
                if active_pd_doc.get("ok") and active_pd_doc.get("raw") is not None:
                    live_pd_doc = active_pd_doc["raw"]
                    result["document_probe"]["live_pddoc"] = _describe_com_object(live_pd_doc)
                    result["document_probe"]["live_checks"] = {
                        "GetNumPages": _safe_call(lambda: getattr(live_pd_doc, "GetNumPages")()),
                        "GetFileName": _safe_call(lambda: getattr(live_pd_doc, "GetFileName")()),
                    }
                    js_object_result = _safe_call(lambda: getattr(live_pd_doc, "GetJSObject")())
                    result["document_probe"]["GetJSObject"] = {
                        key: value for key, value in js_object_result.items() if key != "raw"
                    }
                    if js_object_result.get("ok") and js_object_result.get("raw") is not None:
                        result["document_probe"]["live_js_object"] = _describe_com_object(js_object_result["raw"])
                result["document_probe"]["app_menu_checks"] = {
                    "Preflight": _safe_call(lambda: getattr(app, "MenuItemIsEnabled")("Preflight")),
                    "PrintProduction": _safe_call(lambda: getattr(app, "MenuItemIsEnabled")("PrintProduction")),
                }
                result["document_probe"]["menu_candidate_probe"] = probe_menu_candidates(
                    app,
                    MENU_CANDIDATES,
                    execute_enabled=True,
                )
                if result["workflow_probe"].get("probe_candidates"):
                    result["document_probe"]["workflow_candidate_probe"] = probe_menu_candidates(
                        app,
                        result["workflow_probe"]["probe_candidates"],
                        execute_enabled=True,
                    )
                result["document_probe"]["app_state_before_transitions"] = probe_app_state(app)
                result["document_probe"]["menu_transition_probe"] = {
                    menu_name: probe_menu_transition(app, menu_name)
                    for menu_name in TRANSITION_MENU_CANDIDATES
                }
                if js_object_result.get("ok") and js_object_result.get("raw") is not None:
                    result["document_probe"]["js_candidate_probe"] = probe_js_candidates(
                        js_object_result["raw"],
                        JS_CANDIDATES,
                    )

        for object_name, object_info in result["objects"].items():
            property_checks = object_info.get("property_checks", {})
            for check_name, check_result in list(property_checks.items()):
                if isinstance(check_result, dict) and "raw" in check_result:
                    property_checks[check_name] = {
                        key: value for key, value in check_result.items() if key != "raw"
                    }

        if "document_probe" in result and "live_checks" in result["document_probe"]:
            for check_name, check_result in list(result["document_probe"]["live_checks"].items()):
                if isinstance(check_result, dict) and "raw" in check_result:
                    result["document_probe"]["live_checks"][check_name] = {
                        key: value for key, value in check_result.items() if key != "raw"
                    }
        if "document_probe" in result and "app_menu_checks" in result["document_probe"]:
            for check_name, check_result in list(result["document_probe"]["app_menu_checks"].items()):
                if isinstance(check_result, dict) and "raw" in check_result:
                    result["document_probe"]["app_menu_checks"][check_name] = {
                        key: value for key, value in check_result.items() if key != "raw"
                    }
        if "document_probe" in result and "menu_candidate_probe" in result["document_probe"]:
            for menu_name, menu_info in result["document_probe"]["menu_candidate_probe"].items():
                for field_name, field_result in list(menu_info.items()):
                    if isinstance(field_result, dict) and "raw" in field_result:
                        menu_info[field_name] = {key: value for key, value in field_result.items() if key != "raw"}
        if "document_probe" in result and "workflow_candidate_probe" in result["document_probe"]:
            for menu_name, menu_info in result["document_probe"]["workflow_candidate_probe"].items():
                for field_name, field_result in list(menu_info.items()):
                    if isinstance(field_result, dict) and "raw" in field_result:
                        menu_info[field_name] = {key: value for key, value in field_result.items() if key != "raw"}

        return result
    except Exception as exc:
        return {
            "ok": False,
            "stage": "dispatch",
            "error": f"{type(exc).__name__}: {exc}",
            "traceback": traceback.format_exc(),
        }
    finally:
        try:
            if av_doc is not None:
                _safe_call(lambda: getattr(av_doc, "Close")(1))
        except Exception:
            pass
        try:
            if app is not None:
                _safe_call(lambda: getattr(app, "Exit")())
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def main(argv=None):
    parser = argparse.ArgumentParser(description="Probe Acrobat COM automation surface.")
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON only.")
    parser.add_argument("--pdf", help="Optional PDF path to open and inspect through Acrobat.")
    args = parser.parse_args(argv)

    result = probe_acrobat(pdf_path=args.pdf)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("ok") else 1

    print("== Acrobat COM Probe ==")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
