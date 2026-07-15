"""从 CHANGELOG.md 中提取指定版本的小节，供发布流程注入 Release 正文。

用法::

    python scripts/extract_release_notes.py --version 0.7.1 \
        --changelog CHANGELOG.md --output release_notes.md

若在 CHANGELOG 中找到对应版本小节，则将其内容写入输出文件并打印
``found=true``；若未找到，则写入空文件并打印 ``found=false``，
交由发布流程回退到自动生成的说明。
"""

import argparse
import re
import sys


def extract_release_notes(changelog_text, version):
    """返回 ``version`` 对应的 CHANGELOG 小节正文，找不到时返回 ``None``。

    小节以 ``## [<version>]`` 开头，到下一个 ``## `` 二级标题为止。
    返回值会去掉版本标题行本身（Release 已有版本作为标题），
    仅保留其下的 ``###`` 变更内容并去除首尾空白。
    """
    header = re.compile(r"^## \[" + re.escape(version) + r"\]", re.MULTILINE)
    match = header.search(changelog_text)
    if not match:
        return None

    body_start = changelog_text.find("\n", match.start())
    if body_start == -1:
        return ""
    body_start += 1

    next_header = re.compile(r"^## ", re.MULTILINE).search(changelog_text, body_start)
    body_end = next_header.start() if next_header else len(changelog_text)

    return changelog_text[body_start:body_end].strip()


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", required=True, help="目标版本号，例如 0.7.1")
    parser.add_argument("--changelog", default="CHANGELOG.md", help="CHANGELOG 文件路径")
    parser.add_argument("--output", default="release_notes.md", help="输出的 Release 正文文件路径")
    args = parser.parse_args(argv)

    with open(args.changelog, "r", encoding="utf-8") as handle:
        changelog_text = handle.read()

    notes = extract_release_notes(changelog_text, args.version)
    found = notes is not None

    with open(args.output, "w", encoding="utf-8") as handle:
        handle.write(notes if found else "")

    print(f"found={'true' if found else 'false'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
