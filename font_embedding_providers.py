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
(function () {{
    var profileNames = {profile_names};
    function fail(message) {{
        throw new Error(message);
    }}
    function getProfileByName(name) {{
        if (typeof Preflight !== "undefined" && Preflight.getProfileByName) {{
            try {{ return Preflight.getProfileByName(name); }} catch (e1) {{}}
        }}
        if (typeof preflight !== "undefined" && preflight.getProfileByName) {{
            try {{ return preflight.getProfileByName(name); }} catch (e2) {{}}
        }}
        if (this.preflight && this.preflight.getProfileByName) {{
            try {{ return this.preflight.getProfileByName(name); }} catch (e3) {{}}
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
        fail("未找到可执行的 Acrobat Preflight 字体嵌入 profile。请通过 RATOOLS_ACROBAT_PREFLIGHT_PROFILE 指定本机 profile 名称。");
    }}
    if (typeof this.preflight !== "function") {{
        fail("当前 Acrobat JavaScript 环境未暴露 doc.preflight()，可能不是 Acrobat Pro 或该版本不支持脚本化 Preflight。");
    }}
    var result = this.preflight(profile, true);
    return JSON.stringify({{ ok: true, profile: profileName, result: String(result) }});
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
        pd_doc = None
        try:
            app = win32com_client.Dispatch("AcroExch.App")
            pd_doc = win32com_client.Dispatch("AcroExch.PDDoc")
            if not pd_doc.Open(output_path):
                return FontEmbeddingResult(False, self.provider_id, self.provider_name, "Acrobat 无法打开待处理 PDF")

            js_object = pd_doc.GetJSObject()
            script = self._preflight_script()
            try:
                result = js_object.eval(script)
            except Exception as exc:
                return FontEmbeddingResult(False, self.provider_id, self.provider_name, f"Acrobat Preflight 执行失败：{exc}")

            try:
                pd_doc.Save(1, output_path)
            except Exception as exc:
                return FontEmbeddingResult(False, self.provider_id, self.provider_name, f"Acrobat 保存 PDF 失败：{exc}")

            message = "Acrobat Preflight 已执行"
            if result:
                message = f"{message}：{result}"
            return FontEmbeddingResult(True, self.provider_id, self.provider_name, message)
        finally:
            if pd_doc is not None:
                try:
                    pd_doc.Close()
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
