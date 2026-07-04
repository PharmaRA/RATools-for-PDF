import unittest

from acrobat_probe import (
    build_workflow_probe_candidates,
    build_workflow_registry_index,
    extract_sequence_metadata,
    extract_sequence_import_command,
    find_interesting_members,
    list_public_member_names,
    probe_app_state,
    probe_js_candidates,
    probe_menu_transition,
    probe_menu_candidates,
)


class _FakeComObject:
    Visible = True

    def Open(self):
        return True

    def ExecMenuItem(self):
        return True

    def _private_method(self):
        return False


class _FakeApp:
    def __init__(self):
        self.executed = []
        self.active_tool = "SelectTool"
        self.av_docs = 1
        self.active_doc = "active-doc"

    def MenuItemIsEnabled(self, name):
        return name in {"PrintProduction", "Preflight"}

    def MenuItemExecute(self, name):
        self.executed.append(name)
        if name == "PrintProduction":
            self.active_tool = "PrintProductionTool"
        return f"executed:{name}"

    def GetActiveTool(self):
        return self.active_tool

    def GetNumAVDocs(self):
        return self.av_docs

    def GetActiveDoc(self):
        return self.active_doc


class _FakeJsObject:
    app = "acrobat-js-app"

    def execMenuItem(self, name):
        return f"js-exec:{name}"

    def beginPriv(self):
        return True


class AcrobatProbeHelperTests(unittest.TestCase):
    def test_extract_sequence_import_command_expands_percent_one_placeholder(self):
        command_template = '"C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe" "%1"'

        command = extract_sequence_import_command(command_template, r"E:\tmp\Action04.sequ")

        self.assertEqual(
            command,
            [
                r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                r"E:\tmp\Action04.sequ",
            ],
        )

    def test_build_workflow_registry_index_keeps_title_id_and_relative_path(self):
        entries = [
            {
                "key_name": "c3",
                "values": {
                    "tTitle": "Optimize for Web and Mobile",
                    "aID": "CB7C61DD0538D7707B0D598AB312A0F",
                    "tRelativeDIPath": "Action04.sequ",
                },
            },
            {
                "key_name": "c4",
                "values": {
                    "tTitle": "Optimize Scanned Documents",
                    "aID": "00F5E03036002B85126B8150A9F3434",
                    "tRelativeDIPath": "Action05.sequ",
                },
            },
        ]

        index = build_workflow_registry_index(entries)

        self.assertEqual(index[0]["title"], "Optimize for Web and Mobile")
        self.assertEqual(index[0]["relative_path"], "Action04.sequ")
        self.assertEqual(index[0]["id"], "CB7C61DD0538D7707B0D598AB312A0F")
        self.assertEqual(index[1]["title"], "Optimize Scanned Documents")

    def test_extract_sequence_metadata_reads_workflow_title_and_preflight_profile(self):
        xml_text = """<?xml version='1.0' encoding='UTF-8'?>
<Workflow xmlns='http://ns.adobe.com/acrobat/workflow/2012' title='Optimize for Web and Mobile'>
  <Group>
    <Command name='CALS:Preflight'>
      <Items>
        <Item name='CALS_PREFLIGHT_CMD_PROFILE_DICTKEY' type='text' value='P_7_Embedmissingfonts'/>
        <Item name='CALS_PREFLIGHT_CMD_PROFILE_FINGERPRINT' type='text' value='P9db551f478f00782f340fa57fd08cf08'/>
        <Item name='CALS_PREFLIGHT_CMD_PROFILE_NAME' type='text' value='Embed missing fonts'/>
      </Items>
    </Command>
  </Group>
</Workflow>
"""

        metadata = extract_sequence_metadata(xml_text)

        self.assertEqual(metadata["title"], "Optimize for Web and Mobile")
        self.assertEqual(metadata["preflight_profiles"][0]["name"], "Embed missing fonts")
        self.assertEqual(metadata["preflight_profiles"][0]["dictkey"], "P_7_Embedmissingfonts")
        self.assertEqual(metadata["preflight_profiles"][0]["fingerprint"], "P9db551f478f00782f340fa57fd08cf08")

    def test_build_workflow_probe_candidates_includes_title_action_and_profile_hints(self):
        metadata = {
            "title": "Optimize for Web and Mobile",
            "preflight_profiles": [
                {
                    "name": "Embed missing fonts",
                    "dictkey": "P_7_Embedmissingfonts",
                    "fingerprint": "P9db551f478f00782f340fa57fd08cf08",
                }
            ],
        }

        candidates = build_workflow_probe_candidates(
            metadata,
            workflow_relative_path="Action04.sequ",
            workflow_id="CB7C61DD0538D7707B0D598AB312A0F",
        )

        self.assertIn("Optimize for Web and Mobile", candidates)
        self.assertIn("Action04", candidates)
        self.assertIn("Action04.sequ", candidates)
        self.assertIn("CB7C61DD0538D7707B0D598AB312A0F", candidates)
        self.assertIn("Embed missing fonts", candidates)
        self.assertIn("P_7_Embedmissingfonts", candidates)
        self.assertIn("P9db551f478f00782f340fa57fd08cf08", candidates)

    def test_list_public_member_names_excludes_private_names(self):
        members = list_public_member_names(_FakeComObject())

        self.assertIn("Open", members)
        self.assertIn("ExecMenuItem", members)
        self.assertIn("Visible", members)
        self.assertNotIn("_private_method", members)

    def test_find_interesting_members_matches_preflight_related_keywords(self):
        members = ["Open", "ExecMenuItem", "Save", "PreflightFixup", "PlainProperty"]

        interesting = find_interesting_members(members)

        self.assertEqual(interesting, ["ExecMenuItem", "PreflightFixup", "Save"])

    def test_probe_menu_candidates_checks_enabled_state_and_optional_execute(self):
        app = _FakeApp()

        result = probe_menu_candidates(app, ["PrintProduction", "UnknownMenu"], execute_enabled=True)

        self.assertTrue(result["PrintProduction"]["enabled"]["value"])
        self.assertEqual(result["PrintProduction"]["execute"]["value"], "executed:PrintProduction")
        self.assertFalse(result["UnknownMenu"]["enabled"]["value"])
        self.assertNotIn("execute", result["UnknownMenu"])

    def test_probe_js_candidates_checks_presence_and_callable_invocation(self):
        js_obj = _FakeJsObject()

        result = probe_js_candidates(js_obj, ["app", "execMenuItem", "beginPriv", "missingName"])

        self.assertTrue(result["app"]["exists"])
        self.assertEqual(result["app"]["value"]["value"], "acrobat-js-app")
        self.assertTrue(result["execMenuItem"]["exists"])
        self.assertEqual(result["execMenuItem"]["call_test"]["value"], "js-exec:PrintProduction")
        self.assertTrue(result["beginPriv"]["exists"])
        self.assertEqual(result["beginPriv"]["call_test"]["value"], True)
        self.assertFalse(result["missingName"]["exists"])

    def test_probe_app_state_collects_basic_runtime_snapshot(self):
        app = _FakeApp()

        result = probe_app_state(app)

        self.assertEqual(result["GetActiveTool"]["value"], "SelectTool")
        self.assertEqual(result["GetNumAVDocs"]["value"], 1)
        self.assertEqual(result["GetActiveDoc"]["value"], "active-doc")

    def test_probe_menu_transition_captures_before_and_after_state(self):
        app = _FakeApp()

        result = probe_menu_transition(app, "PrintProduction")

        self.assertEqual(result["before"]["GetActiveTool"]["value"], "SelectTool")
        self.assertEqual(result["execute"]["value"], "executed:PrintProduction")
        self.assertEqual(result["after"]["GetActiveTool"]["value"], "PrintProductionTool")


if __name__ == "__main__":
    unittest.main()
