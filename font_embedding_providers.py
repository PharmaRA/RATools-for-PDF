import json
import os
import platform
import shutil
from dataclasses import dataclass
from pathlib import Path


@dataclass
class FontEmbeddingResult:
    success: bool
    provider_id: str
    provider_name: str
    message: str


class FontEmbeddingProvider:
    provider_id = "base"
    provider_name = "Base provider"

    def availability(self):
        return False, "provider is not implemented"

    def embed_missing_fonts(self, input_path, output_path, font_precheck=None):
        raise NotImplementedError


class AcrobatPreflightFontEmbeddingProvider(FontEmbeddingProvider):
    provider_id = "acrobat_preflight"
    provider_name = "Acrobat Pro Preflight"

    DEFAULT_PROFILE_NAMES = [
        "Embed missing fonts",
        "Embed fonts",
        "Embed all fonts",
        "嵌入缺失的字体",
        "嵌入缺失字体",
        "嵌入字体",
    ]

    def __init__(self):
        self._win32com_client = None

    def _load_win32com(self):
        if self._win32com_client is not None:
            return self._win32com_client
        try:
            import win32com.client  # type: ignore
        except Exception as exc:
            raise RuntimeError("缺少 pywin32，无法调用 Acrobat COM/OLE") from exc
        self._win32com_client = win32com.client
        return self._win32com_client

    def availability(self):
        if platform.system().lower() != "windows":
            return False, "Acrobat COM/OLE 后端仅支持 Windows"
        try:
            win32com_client = self._load_win32com()
            app = win32com_client.Dispatch("AcroExch.App")
            try:
                app.Exit()
            except Exception:
                pass
            return True, "Acrobat COM/OLE 可用"
        except Exception as exc:
            return False, f"未检测到可用的 Acrobat COM/OLE：{exc}"

    def _profile_names(self):
        configured = os.environ.get("RATOOLS_ACROBAT_PREFLIGHT_PROFILE", "").strip()
        if configured:
            return [item.strip() for item in configured.split(";") if item.strip()]
        return list(self.DEFAULT_PROFILE_NAMES)

    def _preflight_script(self):
        script_path = os.environ.get("RATOOLS_ACROBAT_PREFLIGHT_JS", "").strip()
        if script_path:
            path = Path(script_path)
            if not path.exists():
                raise FileNotFoundError(f"Acrobat Preflight JS 文件不存在：{script_path}")
            return path.read_text(encoding="utf-8")

        profile_names = json.dumps(self._profile_names(), ensure_ascii=False)
        return f"""
event.value = (function () {{
try {{
    var currentStep = "init";
    var profileNames = {profile_names};
    var targetDoc = (typeof event !== "undefined" && event && event.target) ? event.target : this;
    function fail(message) {{
        return {{ ok: false, error: String(message), step: currentStep }};
    }}
    function safeGet(obj, key) {{
        try {{
            return obj ? obj[key] : null;
        }} catch (getError) {{
            return null;
        }}
    }}
    function safeCall(obj, methodName, arg1) {{
        try {{
            var method = safeGet(obj, methodName);
            if (typeof method !== "function") {{
                return null;
            }}
            if (arguments.length > 2) {{
                return method.call(obj, arg1);
            }}
            return method.call(obj);
        }} catch (callError) {{
            return null;
        }}
    }}
    function getPreflightApi() {{
        currentStep = "locate Preflight API";
        if (typeof Preflight !== "undefined") {{
            return Preflight;
        }}
        if (typeof preflight !== "undefined") {{
            return preflight;
        }}
        var docPreflight = safeGet(targetDoc, "preflight");
        if (docPreflight) {{
            return docPreflight;
        }}
        return null;
    }}
    function collectProfileNames() {{
        currentStep = "collect installed profiles";
        var names = [];
        var api = getPreflightApi();
        if (!api || typeof safeGet(api, "getNumProfiles") !== "function" || typeof safeGet(api, "getNthProfile") !== "function") {{
            return names;
        }}
        try {{
            var count = safeCall(api, "getNumProfiles");
            for (var idx = 0; idx < count; idx++) {{
                try {{
                    var item = safeCall(api, "getNthProfile", idx);
                    if (!item) {{
                        continue;
                    }}
                    var name = item.name || item.Name || String(item);
                    if (name) {{
                        names.push(String(name));
                    }}
                }} catch (profileError) {{}}
            }}
        }} catch (listError) {{}}
        return names;
    }}
    function getProfileByName(name) {{
        currentStep = "find profile: " + name;
        var api = getPreflightApi();
        var profile = safeCall(api, "getProfileByName", name);
        if (profile) {{
            return profile;
        }}
        return null;
    }}
    var profile = null;
    var profileName = "";
    for (var i = 0; i < profileNames.length; i++) {{
        profile = getProfileByName.call(this, profileNames[i]);
        if (profile) {{
            profileName = profileNames[i];
            break;
        }}
    }}
    if (!profile) {{
        var notFound = fail("未找到可执行的 Acrobat Preflight 字体嵌入 profile。请通过 RATOOLS_ACROBAT_PREFLIGHT_PROFILE 指定本机 profile 名称。");
        notFound.requested_profiles = profileNames;
        notFound.installed_profiles = collectProfileNames();
        return JSON.stringify(notFound);
    }}
    var docPreflight = safeGet(targetDoc, "preflight");
    if (!targetDoc || typeof docPreflight !== "function") {{
        return JSON.stringify(fail("当前 Acrobat JavaScript 环境未暴露 doc.preflight()，可能不是 Acrobat Pro 或该版本不支持脚本化 Preflight。"));
    }}
    currentStep = "execute preflight: " + profileName;
    var thermometer = (typeof app !== "undefined" && app && app.thermometer) ? app.thermometer : null;
    var result = docPreflight.call(targetDoc, profile, false, thermometer);
    return JSON.stringify({{
        ok: true,
        profile: profileName,
        result: String(result),
        numErrors: result && typeof result.numErrors !== "undefined" ? result.numErrors : null,
        numWarnings: result && typeof result.numWarnings !== "undefined" ? result.numWarnings : null,
        numInfos: result && typeof result.numInfos !== "undefined" ? result.numInfos : null,
        numFixed: result && typeof result.numFixed !== "undefined" ? result.numFixed : null,
        numNotFixed: result && typeof result.numNotFixed !== "undefined" ? result.numNotFixed : null
    }});
}} catch (e) {{
    return JSON.stringify({{ ok: false, error: String(e && (e.message || e)), step: currentStep }});
}}
}})();
"""

    def embed_missing_fonts(self, input_path, output_path, font_precheck=None):
        available, reason = self.availability()
        if not available:
            return FontEmbeddingResult(False, self.provider_id, self.provider_name, reason)

        win32com_client = self._load_win32com()
        input_path = os.path.abspath(str(input_path))
        output_path = os.path.abspath(str(output_path))
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        shutil.copy2(input_path, output_path)

        app = None
        av_doc = None
        pd_doc = None
        try:
            app = win32com_client.Dispatch("AcroExch.App")
            av_doc = win32com_client.Dispatch("AcroExch.AVDoc")
            if not av_doc.Open(output_path, ""):
                return FontEmbeddingResult(False, self.provider_id, self.provider_name, "Acrobat 无法打开待处理 PDF")
            pd_doc = av_doc.GetPDDoc()

            script = self._preflight_script()
            try:
                aform_app = win32com_client.Dispatch("AFormAut.App")
                result = aform_app.Fields.ExecuteThisJavascript(script)
            except Exception as exc:
                return FontEmbeddingResult(False, self.provider_id, self.provider_name, f"Acrobat Preflight 执行失败：{exc}")
            if not str(result or "").strip():
                return FontEmbeddingResult(
                    False,
                    self.provider_id,
                    self.provider_name,
                    "Acrobat Preflight 脚本未返回脚本结果，无法确认 profile 是否找到或 fixup 是否执行。"
                    "请确认 Acrobat Pro 支持 Preflight JavaScript，并通过 RATOOLS_ACROBAT_PREFLIGHT_PROFILE 指定本机可用的字体嵌入 profile。",
                )
            try:
                script_result = json.loads(str(result))
            except Exception:
                script_result = None
            if isinstance(script_result, dict) and not script_result.get("ok", False):
                error = script_result.get("error") or str(result)
                details = [f"Acrobat Preflight 脚本报告失败：{error}"]
                step = script_result.get("step")
                if step:
                    details.append(f"失败步骤：{step}")
                requested_profiles = script_result.get("requested_profiles") or []
                installed_profiles = script_result.get("installed_profiles") or []
                if requested_profiles:
                    details.append(f"已尝试 profile：{'; '.join(str(item) for item in requested_profiles)}")
                if installed_profiles:
                    details.append(f"本机可见 profile：{'; '.join(str(item) for item in installed_profiles)}")
                else:
                    details.append("未能从 Acrobat JavaScript API 枚举到本机 profile")
                return FontEmbeddingResult(False, self.provider_id, self.provider_name, "；".join(details))

            try:
                pd_doc.Save(1, output_path)
            except Exception as exc:
                return FontEmbeddingResult(False, self.provider_id, self.provider_name, f"Acrobat 保存 PDF 失败：{exc}")

            message = "Acrobat Preflight 已执行"
            if isinstance(script_result, dict):
                profile = script_result.get("profile") or "未知 profile"
                preflight_result = script_result.get("result") or ""
                message = f"{message}：profile={profile}"
                if preflight_result:
                    message = f"{message}，result={preflight_result}"
                counters = []
                for key in ["numErrors", "numWarnings", "numInfos", "numFixed", "numNotFixed"]:
                    value = script_result.get(key)
                    if value is not None:
                        counters.append(f"{key}={value}")
                if counters:
                    message = f"{message}，{', '.join(counters)}"
            elif result:
                message = f"{message}：{result}"
            return FontEmbeddingResult(True, self.provider_id, self.provider_name, message)
        finally:
            if pd_doc is not None:
                try:
                    pd_doc.Close()
                except Exception:
                    pass
            if av_doc is not None:
                try:
                    av_doc.Close(True)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Exit()
                except Exception:
                    pass


def get_font_embedding_provider(provider_id=None):
    provider_id = (provider_id or os.environ.get("RATOOLS_FONT_EMBED_PROVIDER") or "acrobat_preflight").strip().lower()
    providers = {
        AcrobatPreflightFontEmbeddingProvider.provider_id: AcrobatPreflightFontEmbeddingProvider,
        "acrobat": AcrobatPreflightFontEmbeddingProvider,
    }
    provider_cls = providers.get(provider_id)
    if provider_cls is None:
        raise ValueError(f"未知字体修复 provider：{provider_id}")
    return provider_cls()
