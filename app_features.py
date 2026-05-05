import os


def _is_enabled(value):
    return str(value).strip().lower() not in {"0", "false", "no", "off"}


ENABLE_UPDATE_CHECK = _is_enabled(os.environ.get("RATOOLS_ENABLE_UPDATE_CHECK", "1"))
