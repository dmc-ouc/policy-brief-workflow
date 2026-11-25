"""Read-only local package checks. Web access and real subagents need session checks."""
from __future__ import annotations

import json
import platform
import sys
from pathlib import Path

sys.dont_write_bytecode = True
from build_docx import DEFAULT_TEMPLATE, digest, inspect_template

REQUIRED_FONTS = ("方正小标宋简体", "黑体", "楷体", "仿宋", "宋体")


def font_inventory() -> dict:
    if sys.platform != "win32":
        return {"status": "需由排版工具核对", "required": list(REQUIRED_FONTS)}
    import winreg
    names = []
    errors = []
    subkey = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, subkey) as key:
                for index in range(winreg.QueryInfoKey(key)[1]):
                    names.append(winreg.EnumValue(key, index)[0])
        except FileNotFoundError:
            continue
        except OSError as error:
            errors.append(str(error))
    # Registry display names can be English/localized; no negative font verdict here.
    candidates = {font: [name for name in names if font in name] for font in REQUIRED_FONTS}
    return {
        "status": "仅注册表线索；仍须渲染核对实际字体及替代情况",
        "candidates": candidates,
        "unmatched_names": [font for font, found in candidates.items() if not found],
        "registry_read_errors": errors,
    }


def main() -> int:
    root = Path(__file__).resolve().parents[4]
    relative_files = [
        "AGENTS.md", "README.md",
        ".agents/skills/policy-brief/SKILL.md",
        ".agents/skills/policy-brief/references/sources.md",
        ".agents/skills/policy-brief/references/deliverables.md",
        ".agents/skills/policy-brief/references/layout.md",
        ".agents/skills/policy-brief/references/runtime.md",
        ".agents/skills/policy-brief-polish/SKILL.md",
        ".agents/skills/policy-brief-polish/references/style-v001.md",
        ".agents/skills/policy-brief/assets/report-template.docx",
        ".agents/skills/policy-brief/scripts/build_docx.py",
        ".agents/skills/policy-brief/scripts/export_word.ps1",
    ] + [f".codex/agents/{role}.toml" for role in ("policy-research", "policy-writer", "policy-reviewer")]
    missing = [name for name in relative_files if not (root / name).is_file()]
    errors = []
    config = {}
    if sys.version_info < (3, 11):
        errors.append("需要 Python 3.11 或更新版本，以便验证 TOML；未尝试安装")
    else:
        import tomllib
        for role in ("policy-research", "policy-writer", "policy-reviewer"):
            path = root / ".codex" / "agents" / (role + ".toml")
            try:
                data = tomllib.loads(path.read_text(encoding="utf-8-sig"))
                valid = data.get("name") == role and all(isinstance(data.get(key), str) and data[key].strip() for key in ("description", "developer_instructions"))
                config[role] = valid
                if not valid:
                    errors.append(role + " 缺少有效的角色名称、说明或指令")
            except (OSError, ValueError) as error:
                errors.append(role + ": " + str(error))
    template = {}
    try:
        parts, raw = inspect_template(DEFAULT_TEMPLATE)
        template = {"compatible": True, "sha256": digest(raw), "parts": len(parts)}
    except Exception as error:
        errors.append("模板检查失败：" + str(error))
    report = {
        "local_package_check": "通过" if not missing and not errors else "未通过",
        "python": platform.python_version(),
        "platform": platform.system(),
        "missing_files": missing,
        "errors": errors,
        "agent_config_syntax": config,
        "template": template,
        "fonts": font_inventory(),
        "must_check_in_current_codex_session": [
            "能读取本项目两个技能，且可向任务目录写文件",
            "能真正启动独立子 Agent 并等待返回；TOML 语法通过不代表此能力可用",
            "能搜索并打开原文、识别重定向后的最终域名；搜索摘要不算证据",
            "具备可用 Word 渲染方式并能逐页查看；字体名写入 XML 不证明字体存在",
        ],
        "workflow_ready": "未判定；本脚本不访问网络、不启动子 Agent、不执行 Word，也不启动选题",
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if not missing and not errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
