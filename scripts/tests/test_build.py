Warning: truncated output (original token count: 50394)
Total output lines: 4788

#!/usr/bin/env python3
"""Lightweight tests for scripts/build.py and scripts/shared.py.

Run with: python3 scripts/tests/test_build.py
The harness uses plain assertions and a tiny runner so it has no third-party
dependency (matching the rest of the repo's lean tooling).
"""
from __future__ import annotations

import contextlib
import builtins
import hashlib
import importlib.util
import inspect
import io
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import time
import tracemalloc
import warnings
import zipfile
from pathlib import Path

# Make scripts/ importable when running this file directly.
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from build import (  # noqa: E402
    DIAGRAM_TARGETS,
    HTML_TARGETS,
    PPTX_TARGETS,
    SCREEN_TARGETS,
    main as build_main,
)
from checks import (  # noqa: E402
    _BG_B,
    _BG_G,
    _BG_R,
    _density_bucket,
    _last_content_y,
    _markdown_residue_issues,
    _orphan_last_line,
    _parse_slide_sequence,
    _resume_balance_issues,
    _rhythm_issues,
    check_markdown_residue,
    check_placeholders,
    scan_density,
)
from lint import (  # noqa: E402
    NEGATIVE_EXAMPLE_LINE,
    _blank_block,
    _documented_snippets,
    _emphasis_container_findings,
    _extract_root_vars,
    _off_palette_findings,
    _pair_names,
    _root_token_findings,
    _undefined_token_findings,
    check_all,
    check_cross_template_consistency,
    check_off_palette,
    scan_file,
    scan_text,
)
from optional_deps import MissingDepError, require_pymupdf  # noqa: E402
from shared import (  # noqa: E402
    DIAGRAMS,
    DIAGRAM_TEMPLATES,
    HTML_TEMPLATES,
    MARP_TEMPLATES,
    PARCHMENT_RGB,
    ROOT as REPO_ROOT,
    SCREEN_TEMPLATES,
    TEMPLATES,
    build_targets,
    diagram_targets,
    load_checks_thresholds,
    marp_targets,
    pptx_targets,
    screen_targets,
)
import highlight as highlight_mod  # noqa: E402
import shared as shared_mod  # noqa: E402
import verify as verify_mod  # noqa: E402
from highlight import highlight_code_blocks  # noqa: E402
from site_facts import (  # noqa: E402
    FULL_PUBLIC_FACT_FILES,
    REDIRECT_SITE_FILE,
    check_site_facts,
    site_fact_issues,
    site_structure_issues,
)
from tokens import _mermaid_theme_drift  # noqa: E402
from verify import (  # noqa: E402
    RECOGNIZABLE_FALLBACK_FONT_MARKERS,
    _classify_cjk_font,
    _font_family_key,
)


# --------------------------- helpers ---------------------------

_PASS = 0
_FAIL = 0
_SKIP = 0


def check(name: str, predicate: bool, detail: str = "") -> None:
    global _PASS, _FAIL
    if predicate:
        _PASS += 1
        print(f"OK: {name}")
    else:
        _FAIL += 1
        print(f"ERROR: {name}{(' - ' + detail) if detail else ''}")


def skip(name: str, detail: str = "", *, ci_required: bool = False) -> None:
    """Record an unavailable optional-dependency test without calling it a pass."""
    global _SKIP, _FAIL
    _SKIP += 1
    if ci_required and os.environ.get("CI"):
        _FAIL += 1
        print(f"ERROR: required CI test skipped: {name}{(' - ' + detail) if detail else ''}")
    else:
        print(f"SKIP: {name}{(' - ' + detail) if detail else ''}")


def write_temp_html(body: str, suffix: str = "-en.html") -> Path:
    f = tempfile.NamedTemporaryFile(mode="w", suffix=suffix, delete=False, encoding="utf-8")
    f.write(body)
    f.close()
    return Path(f.name)


def silently(callable_, *args, **kwargs):
    """Run a function with stdout suppressed, return its result."""
    sink = io.StringIO()
    with contextlib.redirect_stdout(sink):
        return callable_(*args, **kwargs)


def run_build_args(args: list[str]) -> tuple[int, str]:
    sink = io.StringIO()
    with contextlib.redirect_stdout(sink):
        rc = build_main(["build.py", *args])
    return rc, sink.getvalue()


def site_fact_file_map() -> dict[str, str]:
    rels = (*FULL_PUBLIC_FACT_FILES, REDIRECT_SITE_FILE)
    return {
        rel: (REPO_ROOT / rel).read_text(encoding="utf-8", errors="replace")
        for rel in rels
    }


# --------------------------- package archive ---------------------------

PACKAGE_MAX_BYTES = 6_000_000
PACKAGE_ROOT_NAME = "kami"
PACKAGE_FORBIDDEN_EXACT = {
    ".claude-plugin/marketplace.json",
    ".gitignore",
    "AGENTS.md",
    "CLAUDE.md",
    "README.md",
    "assets/images/1.png",
    "assets/images/2.png",
    "assets/images/3.png",
    "assets/fonts/TsangerJinKai02-W04.ttf",
    "assets/fonts/TsangerJinKai02-W05.ttf",
    "assets/fonts/SourceHanSerifKR-Regular.otf",
    "assets/fonts/SourceHanSerifKR-Medium.otf",
    "index.html",
    "index-en.html",
    "index-ja.html",
    "index-ko.html",
    "index-tw.html",
    "index-zh.html",
    "llms.txt",
    "robots.txt",
    "scripts/build_metadata.py",
    "scripts/draft-release-notes.py",
    "scripts/package-skill.sh",
    "sitemap.xml",
    "styles.css",
    "vercel.json",
}
PACKAGE_FORBIDDEN_PREFIXES = (
    ".agents/",
    ".claude/",
    ".github/",
    "assets/demos/",
    "assets/examples/",
    "assets/illustrations/",
    "assets/showcase/",
    "plugins/",
    "scripts/tests/",
)
PACKAGE_REQUIRED_ENTRIES = {
    "SKILL.md",
    "CHEATSHEET.md",
    "VERSION",
    "LICENSE",
    "assets/images/logo.svg",
    "assets/fonts/JetBrainsMono.woff2",
    "assets/templates/resume.html",
    "assets/templates/landing-page.html",
    "assets/diagrams/sequence.html",
    "references/design.md",
    "scripts/build.py",
    "scripts/ensure-fonts.sh",
    "scripts/ensure_mathjax.sh",
    "scripts/math_render.py",
    "scripts/mathjax_svg.js",
    "scripts/mathjax-runtime/package.json",
    "scripts/mathjax-runtime/package-lock.json",
    "scripts/site_facts.py",
}


def test_dist_package_contents() -> None:
    archive = REPO_ROOT / "dist" / "kami.zip"
    check("dist/kami.zip exists", archive.exists(), f"missing {archive}")
    if not archive.exists():
        return

    size_bytes = archive.stat().st_size
    check("dist/kami.zip stays below 6MB",
          size_bytes <= PACKAGE_MAX_BYTES,
          f"{size_bytes} bytes > {PACKAGE_MAX_BYTES} bytes")

    with zipfile.ZipFile(archive) as zf:
        names = set(zf.namelist())

    bad_root = sorted(name for name in names if not name.startswith(f"{PACKAGE_ROOT_NAME}/"))
    check("dist/kami.zip uses a Claude-friendly top-level skill folder",
          not bad_root,
          f"entries outside {PACKAGE_ROOT_NAME}/: {', '.join(bad_root)}")

    payload_names = {
        name.removeprefix(f"{PACKAGE_ROOT_NAME}/")
        for name in names
        if name.startswith(f"{PACKAGE_ROOT_NAME}/")
    }
    forbidden = sorted(
        name for name in payload_names
        if name.startswith(PACKAGE_FORBIDDEN_PREFIXES)
        or name in PACKAGE_FORBIDDEN_EXACT
    )
    check("dist/kami.zip excludes site, CI, tests, demos, generated mirrors, and large bundled fonts",
          not forbidden,
          f"forbidden entries: {', '.join(forbidden)}")
    missing_required = sorted(PACKAGE_REQUIRED_ENTRIES - payload_names)
    check("dist/kami.zip keeps required runtime skill files",
          not missing_required,
          f"missing entries: {', '.join(missing_required)}")

    # Structure alone cannot tell a current package from a stale one. The
    # plugin mirror has `build_metadata.py --check`; the ZIP had no equivalent,
    # so editing a source file and forgetting `package-skill.sh` left every
    # check green while the archive Claude Desktop users download stayed on the
    # old content. Compare what is inside against what it was built from.
    stale: list[str] = []
    absent: list[str] = []
    with zipfile.ZipFile(archive) as zf:
        for name in zf.namelist():
            if name.endswith("/"):
                continue
            source = REPO_ROOT / name.removeprefix(f"{PACKAGE_ROOT_NAME}/")
            if not source.exists():
                absent.append(name)
                continue
            if hashlib.sha256(zf.read(name)).digest() != hashlib.sha256(source.read_bytes()).digest():
                stale.append(name)
    check("dist/kami.zip matches the sources it was built from",
          not stale,
          f"{len(stale)} stale entr(ies): {', '.join(sorted(stale)[:5])}"
          " -- run `bash scripts/package-skill.sh`")
    check("dist/kami.zip carries no entry missing from the repo",
          not absent,
          f"entries with no source: {', '.join(sorted(absent)[:5])}")


def test_package_failure_preserves_last_good_archive() -> None:
    """A failed audit must not replace the last archive users can install."""
    script = REPO_ROOT / "scripts" / "package-skill.sh"
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        out = root / "kami.zip"
        out.write_bytes(b"last-good")
        env = dict(os.environ, KAMI_PACKAGE_MAX_BYTES="1")
        result = subprocess.run(
            ["bash", str(script), str(out)],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            env=env,
        )
        leftovers = list(root.glob(".kami-package.*"))
        check("failed package audit preserves the last good archive",
              result.returncode == 1 and out.read_bytes() == b"last-good",
              (result.stdout + result.stderr).strip())
        check("failed package audit removes its candidate directory",
              leftovers == [], str(leftovers))


def test_plugin_metadata_generated() -> None:
    """Claude Code / Codex marketplaces and plugin mirrors must stay generated."""
    script = REPO_ROOT / "scripts" / "build_metadata.py"
    check("build_metadata.py exists", script.exists(), f"missing {script}")
    if not script.exists():
        return

    result = subprocess.run(
        [sys.executable, str(script), "--check"],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
    )
    detail = (result.stdout + result.stderr).strip()
    check("plugin metadata matches generator", result.returncode == 0, detail)


def test_claude_plugin_marketplace_version_matches_version_file() -> None:
    """Claude Code uses this version instead of falling back to a commit hash."""
    version = (REPO_ROOT / "VERSION").read_text(encoding="utf-8").strip()
    marketplace_file = REPO_ROOT / ".claude-plugin" / "marketplace.json"
    check("Claude plugin marketplace metadata exists", marketplace_file.exists())
    if not marketplace_file.exists():
        return

    marketplace = json.loads(marketplace_file.read_text(encoding="utf-8"))
    plugins = marketplace.get("plugins", [])
    kami_plugin = next((plugin for plugin in plugins if plugin.get("name") == "kami"), None)
    check("Claude plugin marketplace includes kami", kami_plugin is not None)
    if not kami_plugin:
        return

    check("Claude plugin marketplace version matches VERSION",
          kami_plugin.get("version") == version,
          f"marketplace={kami_plugin.get('version')!r}, VERSION={version!r}")
    check("Claude plugin marketplace installs the lightweight plugin directory",
          kami_plugin.get("source") == "./plugins/kami",
          f"source={kami_plugin.get('source')!r}")

    plugin_file = REPO_ROOT / "plugins" / "kami" / ".claude-plugin" / "plugin.json"
    check("Claude plugin manifest exists in generated plugin tree", plugin_file.exists())
    if not plugin_file.exists():
        return

    plugin = json.loads(plugin_file.read_text(encoding="utf-8"))
    check("Claude plugin manifest version matches VERSION",
          plugin.get("version") == version,
          f"plugin={plugin.get('version')!r}, VERSION={version!r}")
    check("Claude plugin manifest exposes skills directory",
          plugin.get("skills") == "./skills/",
          f"skills={plugin.get('skills')!r}")


def test_build_metadata_reads_tokens_from_root_argument() -> None:
    from build_metadata import build_codex_plugin, read_token_value

    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        (root / "references").mkdir()
        (root / "references" / "tokens.json").write_text('{"--brand":"#123456"}\n', encoding="utf-8")

        brand_color = read_token_value(root, "brand")
        plugin = build_codex_plugin("9.9.9", brand_color)
        check("build_metadata reads brand token from provided root",
              plugin["interface"]["brandColor"] == "#123456",
              f"brandColor={plugin['interface']['brandColor']}")


def test_catalog_lists_pptx_for_slides() -> None:
    from build_metadata import build_catalog_feed

    catalog = json.loads(build_catalog_feed(REPO_ROOT))
    slide = next(
        entry["item"]
        for entry in catalog["itemListElement"]
        if entry["item"].get("identifier") == "slides"
    )
    pptx_mime = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    check("catalog advertises the editable PPTX slide format",
          pptx_mime in slide.get("encodingFormat", []),
          str(slide.get("encodingFormat")))


# --------------------------- shared registry ---------------------------

def test_registry_consistency() -> None:
    check("HTML_TEMPLATES has 24 entries", len(HTML_TEMPLATES) == 24,
          f"got {len(HTML_TEMPLATES)}")
    check("SCREEN_TARGETS has 3 entries", len(SCREEN_TARGETS) == 3,
          f"got {len(SCREEN_TARGETS)}")
    check("build_targets matches HTML_TEMPLATES key set",
          set(build_targets()) == set(HTML_TEMPLATES))
    check("screen_targets matches SCREEN_TARGETS key set",
          set(screen_targets()) == set(SCREEN_TARGETS))
    check("HTML_TARGETS in build.py matches build_targets()",
          dict(HTML_TARGETS) == build_targets())
    check("DIAGRAM_TARGETS has 18 entries", len(DIAGRAM_TARGETS) == 18,
          f"got {len(DIAGRAM_TARGETS)}")
    check("DIAGRAM_TARGETS in build.py matches shared.diagram_targets()",
          dict(DIAGRAM_TARGETS) == diagram_targets() == dict(DIAGRAM_TEMPLATES))
    check("PPTX_TARGETS has 2 entries", len(PPTX_TARGETS) == 2,
          f"got {len(PPTX_TARGETS)}")
    check("PPTX_TARGETS in build.py matches shared.pptx_targets()",
          dict(PPTX_TARGETS) == pptx_targets())
    registered_sources = {
        **{f"html:{name}": TEMPLATES / spec.source
           for name, spec in HTML_TEMPLATES.items()},
        **{f"screen:{name}": TEMPLATES / source
           for name, source in SCREEN_TEMPLATES.items()},
        **{f"pptx:{name}": TEMPLATES / source
           for name, source in pptx_targets().items()},
        **{f"diagram:{name}": DIAGRAMS / source
           for name, source in DIAGRAM_TEMPLATES.items()},
        **{f"marp:{name}": TEMPLATES / source
           for name, source in MARP_TEMPLATES.items()},
    }
    missing_sources = sorted(
        f"{name} -> {path.relative_to(REPO_ROOT)}"
        for name, path in registered_sources.items()
        if not path.is_file()
    )
    check("every registered template source exists",
          missing_sources == [], ", ".join(missing_sources))
    check("Marp registry maps authoring entries with matching CSS",
          marp_targets() == MARP_TEMPLATES
          and all(
              (TEMPLATES / source).exists()
              and (TEMPLATES / source).with_suffix(".css").exists()
              for source in MARP_TEMPLATES.values()
          ),
          str(MARP_TEMPLATES))
    check("PARCHMENT_RGB is canonical", PARCHMENT_RGB == (0xF5, 0xF4, 0xED))


def test_ci_builds_registered_hard_page_and_pptx_targets() -> None:
    workflow = (REPO_ROOT / ".github" / "workflows" / "check.yml").read_text(
        encoding="utf-8")
    check("CI installs the editable PPTX runtime dependency",
          "python-pptx" in workflow)
    check("CI derives hard-page verification from the shared registry",
          "from shared import HTML_TEMPLATES" in workflow
          and "if spec.build_max_pages" in workflow)
    check("CI verifies both editable PPTX targets",
          "--verify slides\n" in workflow and "--verify slides-en" in workflow)


def test_public_document_kinds_derive_new_registry_entries() -> None:
    from shared import TemplateSpec, public_document_template_kinds

    shared_mod.HTML_TEMPLATES["invoice"] = TemplateSpec("invoice.html", 1)
    try:
        kinds = public_document_template_kinds()
    finally:
        shared_mod.HTML_TEMPLATES.pop("invoice", None)
    check("new registry kinds cannot disappear behind the public-kind allowlist",
          "invoice" in kinds,
          str(sorted(kinds)))


def test_threshold_fallback_includes_resume_balance() -> None:
    original = shared_mod.CHECKS_THRESHOLDS_FILE
    with tempfile.TemporaryDirectory() as d:
        shared_mod.CHECKS_THRESHOLDS_FILE = Path(d) / "missing-thresholds.json"
        shared_mod.load_checks_thresholds.cache_clear()
        try:
            resume = shared_mod.load_checks_thresholds().get("resume_balance")
        finally:
            shared_mod.CHECKS_THRESHOLDS_FILE = original
            shared_mod.load_checks_thresholds.cache_clear()
    check("threshold fallback keeps the resume balance contract",
          resume == {
              "min_fill_pct": 0.83,
              "max_fill_pct": 0.95,
              "max_gap_pct": 0.12,
              "dpi": 36,
          },
          repr(resume))


def test_runner_auto_discovers_tests() -> None:
    names = [name for name, _ in _test_functions()]
    check("test runner auto-discovers Codex update command test",
          "test_check_update_uses_codex_plugin_update_command" in names)
    check("test runner auto-discovers this test",
          "test_runner_auto_discovers_tests" in names)


def test_build_cli_rejects_unexpected_flags() -> None:
    rc, out = run_build_args(["resume", "--verify"])
    check("build.py rejects flags after target",
          rc == 2 and "ERROR: unexpected argument: --verify" in out,
          out.strip())

    rc, out = run_build_args(["--check-density", "-v"])
    check("build.py rejects unknown flags for path-based checks",
          rc == 2 and "ERROR: unexpected argument: -v" in out,
          out.strip())

    rc, out = run_build_args(["--verify", "-v"])
    check("build.py rejects unknown --verify flags",
          rc == 2 and "ERROR: unexpected argument: -v" in out,
          out.strip())

    rc, out = run_build_args(["--check-markdown", "-v"])
    check("build.py rejects unknown --check-markdown flags",
          rc == 2 and "ERROR: unexpected argument: -v" in out,
          out.strip())


def test_long_doc_templates_use_rendered_toc_pages_and_chapter_headers() -> None:
    """Long-doc TOCs must use WeasyPrint target-counter, and running headers
    must follow chapter h1 titles instead of getting stuck on the TOC h2.
    """
    sources = ("long-doc.html", "long-doc-en.html", "long-doc-ko.html")
    required_ids = {
        "#ch-executive-summary",
        "#ch-background",
        "#ch-methodology",
        "#ch-conclusions",
        "#ch-appendix",
    }
    offenders: list[str] = []
    for source in sources:
        text = (TEMPLATES / source).read_text(encoding="utf-8")
        if "target-counter(attr(href), page)" not in text:
            offenders.append(f"{source}: missing target-counter")
        if ".toc-page" in text:
            offenders.append(f"{source}: still has obsolete toc-page wiring")
        missing_ids = sorted(href for href in required_ids if f'href="{href}"' not in text or f'id="{href[1:]}"' not in text)
        if missing_ids:
            offenders.append(f"{source}: missing TOC href/id pairs {missing_ids}")
        h1_block = re.search(r"(?m)^  h1\s*\{(?P<body>.*?)^  \}", text, re.S)
        if not h1_block or "string-set: section-title content();" not in h1_block.group("body"):
            offenders.append(f"{source}: h1 does not set running header")
        h2_block = re.search(r"(?m)^  h2\s*\{(?P<body>.*?)^  \}", text, re.S)
        if h2_block and "string-set:" in h2_block.group("body"):
            offenders.append(f"{source}: h2 still sets running header")

    check("long-doc templates use rendered TOC pages and chapter headers",
          not offenders,
          "; ".join(offenders))


def test_site_facts_repo_clean() -> None:
    rc = silently(check_site_facts, False)
    check("public site facts match shared constants and registries", rc == 0,
          f"check_site_facts returned {rc}")


def test_site_facts_flags_bad_diagram_count() -> None:
    files = site_fact_file_map()
    bad = files["index.html"]
    bad = bad.replace("18 inline SVG diagram types", "17 inline SVG diagram types")
    bad = bad.replace("Eighteen inline SVG diagram types", "Seventeen inline SVG diagram types")
    files["index.html"] = bad

    issues = site_fact_issues(files)
    check("public site facts flag stale diagram counts",
          any("index.html: missing diagram count 18" in issue for issue in issues),
          f"issues: {issues}")


def test_site_facts_cover_developer_install_docs() -> None:
    files = site_fact_file_map()
    command = "npx skills add tw93/kami/plugins/kami -a universal -g -y"
    files["developers.md"] = files["developers.md"].replace(command, "npx skills add stale/path")
    issues = site_fact_issues(files)
    check("public site facts flag stale developer install docs",
          any("developers.md: missing generic agent install command" in issue
              for issue in issues),
          f"issues: {issues}")


def test_site_structure_repo_clean() -> None:
    """Locale pages match index.html's DOM skeleton (redirect script exempt)."""
    issues = site_structure_issues()
    check("locale page structure matches index.html", not issues,
          f"issues: {issues}")


def test_site_structure_flags_locale_drift() -> None:
    files = site_fact_file_map()
    files["index-zh.html"] = files["index-zh.html"].replace(
        '<h2 class="section-title">', '<h3 class="section-title">', 1)

    issues = site_structure_issues(files)
    check("locale structure check flags a drifted heading",
          any("index-zh.html: DOM skeleton drifted" in issue for issue in issues),
          f"issues: {issues}")


def test_chinese_html_templates_keep_single_serif_stack() -> None:
    """Chinese templates must keep --sans pinned to --serif for PDF glyph safety."""
    offenders: list[str] = []
    for name, spec in HTML_TEMPLATES.items():
        source = spec.source
        if name.endswith("-en"):
            continue
        text = (TEMPLATES / source).read_text(encoding="utf-8")
        if "--sans: var(--serif)" not in text and "--sans:  var(--serif)" not in text:
            offenders.append(source)

    check("Chinese HTML templates keep --sans: var(--serif)",
          not offenders,
          f"offenders: {', '.join(offenders)}")


def _ko_stack_offenders(text: str) -> list[str]:
    """Return CSS declarations that reference the bare `"Source Han Serif K"`
    family inside a multi-name fallback stack but omit the real OTF family
    name `"Source Han Serif KR"`.

    The bare name `"Source Han Serif K"` is legitimate on its own only as the
    `@font-face` declared alias (a single-name `font-family: "Source Han Serif K";`
    with no comma, which loads via the file/CDN `src`). Anywhere it appears as a
    fallback item in a comma-separated stack (`--serif`, `--mono`, `@page`
    margin boxes, `code`/`pre`, ...), `"Source Han Serif KR"` MUST sit alongside
    it, or an offline Linux skill install cannot resolve the
    ensure-fonts.sh-downloaded font by name.

    Detection: scan only `font-family` / `--serif` / `--sans` / `--mono`
    declaration values (up to the next `;`, never crossing `{`/`}`). The token
    `"Source Han Serif K"` (closing quote after `K`) never matches
    `"Source Han Serif KR"`, so a value that contains the bare token AND a comma
    (i.e. a fallback stack, not a bare `@font-face` alias) must also contain KR.
    """
    bare = '"Source Han Serif K"'
    kr = '"Source Han Serif KR"'
    decl_re = re.compile(r"(?:font-family|--serif|--sans|--mono)\s*:\s*([^;{}]*)", re.IGNORECASE)
    offenders: list[str] = []
    for m in decl_re.finditer(text):
        value = m.group(1)
        if bare in value and "," in value and kr not in value:
            offenders.append(" ".join(value.split()))
    return offenders


def test_korean_templates_carry_resolvable_serif_name() -> None:
    """Every KO fallback stack that names `Source Han Serif K` must also name
    `Source Han Serif KR` (the actual family of the bundled OTFs), so the font
    resolves by name on an offline Linux skill install. Checks per-declaration,
    not just per-file, so a complete `--serif` cannot mask an incomplete local
    stack (page-margin header/footer, code/pre, mono).
    """
    offenders: list[str] = []
    ko_sources = [spec.source for name, spec in HTML_TEMPLATES.items() if name.endswith("-ko")]
    ko_sources += [source for name, source in SCREEN_TEMPLATES.items() if name.endswith("-ko")]
    # Guard against vacuous green: with zero -ko templates the offender loop
    # never runs and the check below would pass while enforcing nothing.
    check("Korean template set is non-empty", bool(ko_sources),
          "no -ko templates found in the registries")
    for source in ko_sources:
        text = (TEMPLATES / source).read_text(encoding="utf-8")
        for bad in _ko_stack_offenders(text):
            offenders.append(f"{source}: {bad}")

    check("Korean fallback stacks all carry Source Han Serif KR",
          not offenders,
          f"offenders: {'; '.join(offenders)}")


# ---------- sibling placeholder parity (issue #38 class) ----------

# Repeated template structures whose placeholder hints must repeat the first
# block verbatim. Hint richness degrading from block 1 to later siblings makes
# fillers (human or agent) produce degraded copy; see issue #38. Cycle length N
# means placeholders repeat in groups of N (e.g. Role/Actions/Impact rows).
_SIBLING_PARITY_SPECS = (
    ("resume*.html", r'class="proj-text">(\{\{.*?\}\})', 3),
    ("resume*.html", r'class="proj-role">(\{\{.*?\}\})', 1),
    ("resume*.html", r'class="conv-body">\s*(\{\{.*?\}\})', 1),
    ("resume*.html", r'class="os-desc">(\{\{.*?\}\})', 1),
    ("resume*.html", r'class="art-stats">(\{\{.*?\}\})', 1),
    ("portfolio*.html", r'class="project-block">\s*<h3>[^<]*</h3>\s*<p>(\{\{.*?\}\})</p>', 3),
    ("portfolio*.html", r'class="project-type">(\{\{.*?\}\})', 1),
    ("portfolio*.html", r'class="project-date">(\{\{.*?\}\})', 1),
    ("one-pager*.html", r'<li>(\{\{(?:短 bullet|Short bullet|짧은 bullet).*?\}\})</li>', 3),
    ("long-doc*.html", r'(\{\{(?:一段论述|A paragraph|한 단락 논술).*?\}\})', 1),
)


def test_sibling_placeholder_hints_stay_in_parity() -> None:
    """Same-structure sibling blocks must carry identical placeholder hints."""
    matched = 0
    offenders: list[str] = []
    for glob_pattern, regex, cycle in _SIBLING_PARITY_SPECS:
        rx = re.compile(regex, re.DOTALL)
        for path in sorted(TEMPLATES.glob(glob_pattern)):
            hits = rx.findall(path.read_text(encoding="utf-8"))
            if not hits:
                offenders.append(f"{path.name}: no match for {regex[:40]!r} (stale spec?)")
                continue
            matched += 1
            if len(hits) % cycle != 0:
                offenders.append(f"{path.name}: {len(hits)} hint(s) not divisible by cycle {cycle}")
                continue
            first = hits[:cycle]
            for start in range(cycle, len(hits), cycle):
                block = hits[start:start + cycle]
                if block != first:
                    offenders.append(
                        f"{path.name}: block {start // cycle + 1} diverges from block 1: "
                        f"{block} != {first}")
    check("sibling parity specs matched across template families", matched >= 30,
          f"only {matched} template/spec matches; spec table may be stale")
    check("repeated blocks carry identical placeholder hints", not offenders,
          "; ".join(offenders[:6]))


def test_font_fallback_markers_recognize_pt_serif() -> None:
    """macOS without Charter may render English fallbacks as PT Serif."""
    embedded = {"DROIWJ+PT-Serif", "ZBEAAE+JetBrains-Mono"}
    fallback_present = any(
        marker in font for font in embedded
        for marker in RECOGNIZABLE_FALLBACK_FONT_MARKERS
    )
    check("font fallback markers recognize PT-Serif",
          fallback_present,
          f"markers: {RECOGNIZABLE_FALLBACK_FONT_MARKERS}")


def test_font_probe_rejects_empty_and_truncated_bundles() -> None:
    import optional_deps as optional_deps_mod

    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        font_dir = root / "assets" / "fonts"
        font_dir.mkdir(parents=True)
        (font_dir / "Empty.ttf").write_bytes(b"")
        (font_dir / "Truncated.otf").write_bytes(
            b"OTTO\x00\x01\x00\x00\x00\x00\x00\x00"
        )

        original_root = optional_deps_mod.ROOT
        original_which = optional_deps_mod.shutil.which
        try:
            optional_deps_mod.ROOT = root
            optional_deps_mod.shutil.which = lambda _name: None
            report = optional_deps_mod._probe_font(
                "Broken Font",
                ("Empty.ttf", "Truncated.otf"),
                "negative-control font",
            )
        finally:
            optional_deps_mod.ROOT = original_root
            optional_deps_mod.shutil.which = original_which

    check("font probe rejects empty and truncated bundled files",
          report["status"] == "degraded"
          and report["bundled"] == []
          and "Empty.ttf" in report["detail"]
          and "Truncated.otf" in report["detail"],
          str(report))


def test_print_surfaces_have_no_ornamental_brand_lines() -> None:
    """Print hierarchy comes from type, spacing, labels, and fill, not ticks."""
    patterns = {
        "brand side rule": re.compile(
            r"border-left:\s*[\d.]+(?:pt|px)\s+solid\s+var\(--brand\)"
        ),
        "eyebrow tick": re.compile(
            r"\.(?:eyebrow|ticker-eyebrow|cover-eyebrow)::before"
        ),
        "short cover or contact rule": re.compile(
            r"(?:class=[\"'](?:cover-line|contact-line)[\"']|"
            r"\.(?:cover-line|contact-line)\s*\{)"
        ),
    }
    offenders: list[str] = []
    sources = list(TEMPLATES.glob("*.html")) + list((REPO_ROOT / "assets" / "demos").glob("*.html"))
    for path in sorted(sources):
        text = path.read_text(encoding="utf-8")
        for label, pattern in patterns.items():
            if pattern.search(text):
                offenders.append(f"{path.name}: {label}")

    for name in ("slides.py", "slides-en.py"):
        text = (TEMPLATES / name).read_text(encoding="utf-8")
        if "def add_line(" in text:
            offenders.append(f"{name}: generic decorative add_line helper")

    check("print surfaces contain no ornamental brand lines",
          not offenders,
          f"offenders: {', '.join(offenders)}")


def test_shipped_surfaces_keep_subtractive_defaults() -> None:
    """A default surface stays flat, small-radius, and free of filler motion."""
    offenders: list[str] = []

    print_sources = [
        path for path in TEMPLATES.glob("*.html")
        if not path.name.startswith("landing-page")
    ] + list((REPO_ROOT / "assets" / "demos").glob("*.html"))
    large_print_radius = re.compile(
        r"border-radius:\s*(?:[789]|[1-9][0-9])(?:\.[0-9]+)?pt"
    )
    for path in sorted(print_sources):
        if large_print_radius.search(path.read_text(encoding="utf-8")):
            offenders.append(f"{path.name}: screen-sized print radius")

    landing_forbidden = (
        "filter: blur",
        "linear-gradient(",
        ".gallery-frame::after",
    )
    default_surface_selectors = (
        ".hero",
        ".section-head",
        ".demo-card",
        ".price-card",
        ".gallery-caption",
    )
    for path in sorted(TEMPLATES.glob("landing-page*.html")):
        text = path.read_text(encoding="utf-8")
        for token in landing_forbidden:
            if token in text:
                offenders.append(f"{path.name}: {token}")
        for selector in default_surface_selectors:
            match = re.search(re.escape(selector) + r"\s*\{([^}]*)\}", text, re.DOTALL)
            if match and "box-shadow" in match.group(1):
                offenders.append(f"{path.name}: {selector} shadow")

    public_pages = [
        REPO_ROOT / "index.html",
        REPO_ROOT / "index-zh.html",
        REPO_ROOT / "index-tw.html",
        REPO_ROOT / "index-ja.html",
        REPO_ROOT / "index-ko.html",
    ]
    public_forbidden = (
        "border-left: 1.4pt solid var(--brand)",
        "box-shadow: 0 4px 24px",
        'class="dash demo"',
        'class="tag brush"',
        'class="shadow-row"',
    )
    for path in public_pages:
        text = path.read_text(encoding="utf-8")
        for token in public_forbidden:
            if token in text:
                offenders.append(f"{path.name}: {token}")

    site_css = (REPO_ROOT / "styles.css").read_text(encoding="utf-8")
    for token in ("@keyframes fadeIn", "ul.dash", ".tag.brush", ".shadow-row"):
        if token in site_css:
            offenders.append(f"styles.css: {token}")

    for selector in (".family", ".comp", ".chart-card", ".quote"):
        match = re.search(re.escape(selector) + r"\s*\{([^}]*)\}", site_css, re.DOTALL)
        block = match.group(1) if match else ""
        if "box-shadow" in block:
            offenders.append(f"styles.css: {selector} shadow")
        if selector in {".family", ".comp", ".chart-card"} and re.search(
                r"\bborder:\s*1px", block):
            offenders.append(f"styles.css: {selector} stacked border")
        if selector == ".quote" and "border-left" in block:
            offenders.append("styles.css: quote side rule")

    check("shipped surfaces keep subtractive visual defaults",
          not offenders,
          f"offenders: {', '.join(offenders)}")


def test_print_radius_guidance_matches_shipped_range() -> None:
    """The quick reference must describe the same restrained range as templates."""
    cheatsheet = (REPO_ROOT / "CHEATSHEET.md").read_text(encoding="utf-8")
    design = (REPO_ROOT / "references" / "design.md").read_text(encoding="utf-8")
    check("print radius guidance matches shipped 2-6pt range",
          "within `2-6pt`" in cheatsheet
          and "within 2-6pt" in design
          and "Do not invent intermediate steps" not in cheatsheet,
          "radius guidance drifted from shipped template values")


def test_public_site_typography_contract_matches_templates() -> None:
    """Public prose must teach the one-serif contract templates actually ship."""
    pages = [
        REPO_ROOT / "index.html",
        REPO_ROOT / "index-zh.html",
        REPO_ROOT / "index-tw.html",
        REPO_ROOT / "index-ja.html",
        REPO_ROOT / "index-ko.html",
    ]
    stale = (
        "Chinese uses serif headlines and sans body",
        "中文标题用 serif、正文用 sans",
        "中文標題用 serif、正文用 sans",
        "중문은 제목에 세리프, 본문에 산세리프",
    )
    offenders = [
        path.name for path in pages
        if any(token in path.read_text(encoding="utf-8") for token in stale)
    ]
    design = (REPO_ROOT / "references" / "design.md").read_text(encoding="utf-8")
    check("public typography contract matches one-serif templates",
          not offenders
          and "One serif family per page for headlines and body" in design,
          f"offenders: {', '.join(offenders)}")


def test_landing_page_ctas_stack_at_320px() -> None:
    """Localized double CTAs must fit the smallest supported viewport."""
    offenders = []
    for path in sorted(TEMPLATES.glob("landing-page*.html")):
        text = path.read_text(encoding="utf-8")
        if not (
            "@media (max-width: 360px)" in text
            and ".hero-cta { flex-direction: column; align-items: stretch" in text
            and ".btn-ghost { width: 100%; }" in text
        ):
            offenders.append(path.name)
    check("landing page CTAs stack at 320px",
          not offenders,
          f"offenders: {', '.join(offenders)}")


def test_public_site_og_dimensions_match_showcase_image() -> None:
    """Social metadata must describe the image bytes platforms will fetch."""
    image = (REPO_ROOT / "assets" / "showcase" / "kami-landing.png").read_bytes()
    valid_png = image[:8] == b"\x89PNG\r\n\x1a\n" and image[12:16] == b"IHDR"
    width = int.from_bytes(image[16:20], "big") if valid_png else 0
    height = int.from_bytes(image[20:24], "big") if valid_png else 0
    offenders = []
    for name in ("index.html", "index-zh.html", "index-tw.html", "index-ja.html", "index-ko.html"):
        text = (REPO_ROOT / name).read_text(encoding="utf-8")
        declared_width = re.search(r'og:image:width" content="(\d+)"', text)
        declared_height = re.search(r'og:image:height" content="(\d+)"', text)
        if not (
            declared_width and int(declared_width.group(1)) == width
            and declared_height and int(declared_height.group(1)) == height
        ):
            offenders.append(name)
    check("public OG dimensions match the showcase PNG",
          valid_png and not offenders,
          f"image={width}x{height} offenders={', '.join(offenders)}")


def test_public_site_teaches_registered_tints_and_exact_radii() -> None:
    """Visible examples must describe tokens, not obsolete alpha recipes."""
    pages = [
        REPO_ROOT / "index.html",
        REPO_ROOT / "index-zh.html",
        REPO_ROOT / "index-tw.html",
        REPO_ROOT / "index-ja.html",
        REPO_ROOT / "index-ko.html",
    ]
    stale_tint_labels = (
        'class="opacity">0.18',
        'class="tag calm">Light 0.08',
        'class="tag standard">Standard 0.18',
        'class="tag calm">极淡 0.08',
        'class="tag standard">标准 0.18',
        'class="tag calm">極淡 0.08',
        'class="tag standard">標準 0.18',
        "equivalent solid hex",
        "等效实色",
        "等效實色",
        "등가 솔리드 hex",
    )
    offenders: list[str] = []
    for path in pages:
        text = path.read_text(encoding="utf-8")
        for token in stale_tint_labels:
            if token in text:
                offenders.append(f"{path.name}: {token}")
        for radius in ("2", "4"):
            if f'class="box" style="border-radius:{radius}px"' in text:
                offenders.append(f"{path.name}: {radius}px print radius")
            if f'class="box" style="border-radius:{radius}pt"' not in text:
                offenders.append(f"{path.name}: missing {radius}pt print radius")

    check("public site teaches registered tints and exact print radii",
          not offenders,
          f"offenders: {', '.join(offenders)}")


def test_reviewed_demo_details_keep_quiet_hierarchy() -> None:
    """Small editorial cues should not turn back into colored or framed UI."""
    def css_block(text: str, selector: str) -> str:
        match = re.search(re.escape(selector) + r"\s*\{([^}]*)\}", text, re.DOTALL)
        return match.group(1) if match else ""

    offenders: list[str] = []
    resume_surfaces = [
        TEMPLATES / "resume.html",
        TEMPLATES / "resume-en.html",
        TEMPLATES / "resume-ko.html",
        REPO_ROOT / "assets" / "demos" / "demo-musk-resume.html",
        REPO_ROOT / "assets" / "demos" / "demo-resume-ko.html",
    ]
    for path in resume_surfaces:
        text = path.read_text(encoding="utf-8")
        highlight = css_block(text, ".os-highlight")
        label = css_block(text, ".os-highlight .tag")
        if "background: var(--ivory)" not in highlight:
            offenders.append(f"{path.name}: chromatic highlight fill")
        if "background: transparent…30394 tokens truncated…   fail_closed=True,
    )
    check("unresolved functional visibility is excluded and marked ambiguous",
          ambiguous_state and "AMBIGUOUS-VISIBILITY" not in ambiguous_text,
          f"ambiguous={ambiguous_state} text={ambiguous_text!r}")

    non_rendered_html = visible_html_text(
        '<svg><title>SVG-TITLE</title><desc>SVG-DESC</desc></svg>'
        '<svg><defs><text>SVG-DEFS</text></defs>'
        '<symbol><text>SVG-SYMBOL</text></symbol>'
        '<metadata>SVG-METADATA</metadata>'
        '<clipPath><text>SVG-CLIP</text></clipPath>'
        '<mask><text>SVG-MASK</text></mask>'
        '<pattern><text>SVG-PATTERN</text></pattern>'
        '<text>SVG-RENDERED</text></svg>'
        '<noembed>NOEMBED-TEXT</noembed>'
        '<noframes>NOFRAMES-TEXT</noframes>'
        '<datalist><option>DATALIST-OPTION</option></datalist>'
        '<ruby>base<rp>RUBY-FALLBACK</rp><rt>annotation</rt></ruby>'
        '<p>RENDERED-TEXT</p>',
        fail_closed=True,
    )
    check("visible text skips non-rendered metadata and fallback containers",
          all(marker not in non_rendered_html for marker in (
              "SVG-TITLE", "SVG-DESC", "SVG-DEFS", "SVG-SYMBOL",
              "SVG-METADATA", "SVG-CLIP", "SVG-MASK", "SVG-PATTERN",
              "NOEMBED-TEXT", "NOFRAMES-TEXT", "DATALIST-OPTION",
              "RUBY-FALLBACK",
          ))
          and "SVG-RENDERED" not in non_rendered_html
          and "RENDERED-TEXT" in non_rendered_html,
          repr(non_rendered_html))

    hidden_svg_text = visible_html_text(
        '<svg><text display="none">DISPLAY-SECRET</text>'
        '<text visibility="hidden">VISIBILITY-SECRET</text>'
        '<text opacity="0">OPACITY-SECRET</text>'
        '<text>SVG-VISIBLE</text></svg>',
        fail_closed=True,
    )
    check("visible text respects SVG presentation attributes",
          all(marker not in hidden_svg_text for marker in (
              "DISPLAY-SECRET", "VISIBILITY-SECRET", "OPACITY-SECRET",
          ))
          and "SVG-VISIBLE" not in hidden_svg_text,
          repr(hidden_svg_text))

    off_viewport_svg_text = visible_html_text(
        '<svg viewBox="0 0 100 100">'
        '<text x="-99999" y="20">OFF-LEFT</text>'
        '<text x="10000" y="20">OFF-RIGHT</text>'
        '<text x="101" y="20">JUST-OFF-RIGHT</text>'
        '<text x="-99" y="20">JUST-OFF-LEFT</text>'
        '<text x="101%" y="20">PERCENT-OFF-RIGHT</text>'
        '<text x="99" y="20" dx="10">DELTA-OFF-RIGHT</text>'
        '<svg x="101" viewBox="0 0 10 10"><text x="0" y="5">NESTED-OFF</text></svg>'
        '<text x="20" y="20">SVG-IN-VIEW</text></svg>',
        fail_closed=True,
    )
    check("visible text rejects SVG coordinates outside the viewport",
          "OFF-LEFT" not in off_viewport_svg_text
          and "OFF-RIGHT" not in off_viewport_svg_text
          and "JUST-OFF-RIGHT" not in off_viewport_svg_text
          and "JUST-OFF-LEFT" not in off_viewport_svg_text
          and "PERCENT-OFF-RIGHT" not in off_viewport_svg_text
          and "DELTA-OFF-RIGHT" not in off_viewport_svg_text
          and "NESTED-OFF" not in off_viewport_svg_text
          and "SVG-IN-VIEW" not in off_viewport_svg_text,
          repr(off_viewport_svg_text))

    malformed_table = (
        '<div hidden><table></div>MALFORMED-SECRET</table></div>'
        '<p>MALFORMED-VISIBLE</p>'
    )
    malformed_coverage_text = visible_html_text(malformed_table, fail_closed=True)
    malformed_residue_text = visible_html_text(malformed_table)
    check("crossed HTML closes split coverage and residue conservatively",
          "MALFORMED-SECRET" not in malformed_coverage_text
          and "MALFORMED-SECRET" in malformed_residue_text,
          f"coverage={malformed_coverage_text!r} residue={malformed_residue_text!r}")

    optional_end_html = visible_html_text(
        '<p>PARAGRAPH-FIRST<div>PARAGRAPH-SECOND</div><p>PARAGRAPH-THIRD'
        '<ul><li>LIST-FIRST<li>LIST-SECOND</ul>'
        '<table><tr><td>CELL-FIRST<td>CELL-SECOND</tr></table>',
        fail_closed=True,
    )
    check("standard optional HTML end tags preserve visible content",
          all(marker in optional_end_html for marker in (
              "PARAGRAPH-FIRST", "PARAGRAPH-SECOND", "PARAGRAPH-THIRD",
              "LIST-FIRST", "LIST-SECOND", "CELL-FIRST", "CELL-SECOND",
          )),
          repr(optional_end_html))

    self_closing_svg_text, self_closing_svg_ambiguous = visible_html_evidence(
        '<svg viewBox="0 0 100 100"><path d="M0 0L10 10" /></svg>'
        '<p>VISIBLE-AFTER-SVG</p>',
        fail_closed=True,
    )
    check("SVG foreign-content self-closing tags close without tainting HTML",
          self_closing_svg_text.strip() == "VISIBLE-AFTER-SVG"
          and not self_closing_svg_ambiguous,
          f"text={self_closing_svg_text!r} "
          f"ambiguous={self_closing_svg_ambiguous}")


def test_visibility_resolves_document_custom_properties() -> None:
    """Resolve only unconditional custom properties inherited from the root."""
    from checks import visible_html_evidence
    from content import html_resource_evidence

    visible_text, visible_ambiguous = visible_html_evidence(
        "<style>:root { --ink: #504e49 } body { color: var(--ink) }</style>"
        "<body>VISIBLE</body>",
        fail_closed=True,
    )
    check("document custom property makes body text deterministic",
          "VISIBLE" in visible_text and not visible_ambiguous,
          f"text={visible_text!r} ambiguous={visible_ambiguous}")

    template_results = []
    for path in sorted(TEMPLATES.glob("*.html")):
        raw = path.read_text(encoding="utf-8")
        text, ambiguous = visible_html_evidence(raw, fail_closed=True)
        template_results.append((path.name, bool(text.strip()), ambiguous, "<svg" in raw.lower()))
    check("all shipped HTML templates expose visible text under fail-closed checks",
          bool(template_results) and all(result[1] for result in template_results),
          str([result for result in template_results if not result[1]]))
    check("templates without SVG have deterministic visible text",
          all(ambiguous is False for _, _, ambiguous, has_svg in template_results
              if not has_svg),
          str([result for result in template_results if not result[3] and result[2]]))

    ambiguous_cases = [
        '<style>:root{--ink:#000}.dark{--ink:transparent}p{color:var(--ink)}</style>'
        '<div class="dark"><p>SECRET</p></div>',
        '<style>:root{--ink:#000}@media print{:root{--ink:transparent}}'
        'p{color:var(--ink)}</style><p>SECRET</p>',
        '<style>@property --ink{syntax:"<color>";inherits:false;initial-value:transparent}'
        ':root{--ink:#000}p{color:var(--ink)}</style><p>SECRET</p>',
        '<style>:root{--ink:#000}p{color:var(--ink)}</style>'
        '<p style="--ink:transparent">SECRET</p>',
        '<style>:root{--ink:#000}p{color:var(--ink)}</style>'
        '<p style=--ink:transparent>SECRET</p>',
        '<style>:root{--ink:#000}p{color:var(--ink)}</style>'
        "<p style='--ink:transparent'>SECRET</p>",
        '<style>:root{--ink:#000}p{color:var(--ink)}</style>'
        '<p style="--ink:trans&#112;arent">SECRET</p>',
    ]
    ambiguous_results = [
        visible_html_evidence(raw, fail_closed=True) for raw in ambiguous_cases
    ]
    check("conflicting or registered custom properties remain fail-closed",
          all("SECRET" not in text and ambiguous
              for text, ambiguous in ambiguous_results),
          repr(ambiguous_results))

    scoped_cases = [
        '<style>.theme{--d:block}p{display:var(--d,none)}</style>'
        '<div class="theme"></div><p>SECRET</p>',
        '<style>@media screen{:root{--d:block}}p{display:var(--d,none)}</style>'
        '<p>SECRET</p>',
        '<style>p{display:var(--d,none)}</style>'
        '<div style="--d:block"></div><p>SECRET</p>',
        '<style>.noop{--payload:"x; --ink:#000"}'
        'p{color:var(--ink,transparent)}</style><p>SECRET</p>',
        '<style>:root{--d:block}.hide{--junk:{x};--d:none}'
        'p{display:var(--d)}</style><p class="hide">SECRET</p>',
        '<style>:root{--INK:#000}p{color:var(--ink,transparent)}</style>'
        '<p>SECRET</p>',
    ]
    scoped_results = [
        visible_html_evidence(raw, fail_closed=True) for raw in scoped_cases
    ]
    check("scoped and malformed custom properties cannot expose hidden text",
          all("SECRET" not in text for text, _ in scoped_results),
          repr(scoped_results))

    scoped_assets, _ = html_resource_evidence(
        '<style>.theme{--d:block}img{display:var(--d,none)}</style>'
        '<div class="theme"></div><img src="required.svg">'
    )
    check("scoped custom properties cannot expose hidden resources",
          "required.svg" not in scoped_assets, repr(scoped_assets))

    hidden_cases = [
        '<style>:root{--ink:transparent}p{color:var(--ink)}</style><p>SECRET</p>',
        '<style>:root{--d:none}p{display:var(--d, block)}</style><p>SECRET</p>',
        '<p style=\'display:none;--x:";display:block"\'>SECRET</p>',
        '<p style="--HIDE:none;--hide:block;display:var(--HIDE)">SECRET</p>',
    ]
    hidden_results = [
        visible_html_evidence(raw, fail_closed=True) for raw in hidden_cases
    ]
    check("resolved hiding custom properties exclude their text deterministically",
          all("SECRET" not in text and not ambiguous
              for text, ambiguous in hidden_results),
          repr(hidden_results))

    benign_cases = [
        '<style>:root{--ink:#000}html{--ink:#000}p{color:var(--ink)}</style>'
        '<p>SHOWN</p>',
        '<style>:root{--base:#111;--ink:var(--base)}p{color:var(--ink)}</style>'
        '<p>SHOWN</p>',
        '<style>:root{--ink:#111}</style><p style="color:var(--ink)">SHOWN</p>',
    ]
    benign_results = [
        visible_html_evidence(raw, fail_closed=True) for raw in benign_cases
    ]
    check("identical and nested custom properties remain visible",
          all("SHOWN" in text and not ambiguous
              for text, ambiguous in benign_results),
          repr(benign_results))


def test_coverage_checks_asset_attributes() -> None:
    from checks import visible_html_text
    from content import (
        coverage_issues,
        html_resource_attributes,
        html_resource_evidence,
    )

    raw = (
        '<img src="./images/product-shot.png" alt="Product">'
        '<img src="images/product-shot@2x.webp" alt="Product at high density">'
        '<template><img src="hidden-shot.png"></template>'
        '<a href="linked-only.png">not embedded</a>'
    )
    attrs = html_resource_attributes(raw)
    present, checked, _ = coverage_issues(
        {"image": "product-shot.png", "images": ["product-shot@2x.webp"]}, "", attrs
    )
    missing, _, _ = coverage_issues({"image": "missing-shot.png"}, "", attrs)
    check("coverage accepts image paths present in src and srcset",
          present == [] and checked == 2, f"issues={present} attrs={attrs}")

    hidden_svg_attrs = html_resource_attributes(
        '<svg><defs><image href="hidden-def.png"></image></defs>'
        '<symbol><image href="hidden-symbol.png"></image></symbol>'
        '<image x="0" y="0" width="10" height="10" '
        'href="visible-svg.png"></image></svg>'
    )
    check("asset coverage skips SVG definition resources",
          hidden_svg_attrs == {"visible-svg.png"}, repr(hidden_svg_attrs))

    hidden_asset_cases = [
        '<svg><image display="none" href="required.svg"></image></svg>',
        '<svg><image visibility="hidden" href="required.svg"></image></svg>',
        '<svg><image opacity="0" href="required.svg"></image></svg>',
        '<img style="width:0;height:0" src="required.svg">',
        '<img width="0" height="0" src="required.svg">',
        '<svg width="0" height="0"><image href="required.svg"></image></svg>',
        '<img style="position:absolute;left:-99999px" src="required.svg">',
        '<img style="transform:matrix(0,0,0,0,0,0)" src="required.svg">',
        '<img style="width:calc(0px);height:calc(0px)" src="required.svg">',
        '<svg viewBox="0 0 100 100"><image x="-99999" href="required.svg"></image></svg>',
        '<svg viewBox="0 0 100 100"><image x="10000" href="required.svg"></image></svg>',
        '<svg viewBox="0 0 100 100"><image x="101" y="0" width="10" height="10" href="required.svg"></image></svg>',
        '<svg viewBox="0 0 100 100"><image x="101%" y="0" width="10" height="10" href="required.svg"></image></svg>',
        '<svg viewBox="0 0 100 100"><svg x="101"><image x="0" y="0" width="10" height="10" href="required.svg"></image></svg></svg>',
        '<svg viewBox="0 0 100 100"><defs><clipPath id="empty"></clipPath></defs>'
        '<image x="0" y="0" width="20" height="20" '
        'clip-path="url(#empty)" href="required.svg"></image></svg>',
        '<svg viewBox="0 0 100 100"><defs><mask id="empty"></mask></defs>'
        '<image x="0" y="0" width="20" height="20" '
        'mask="url(#empty)" href="required.svg"></image></svg>',
        '<svg viewBox="0 0 100 100"><defs><filter id="empty"></filter></defs>'
        '<image x="0" y="0" width="20" height="20" '
        'filter="url(#empty)" href="required.svg"></image></svg>',
    ]
    check("asset coverage rejects non-rendered CSS and presentation forms",
          all(not html_resource_attributes(case) for case in hidden_asset_cases),
          str([html_resource_attributes(case) for case in hidden_asset_cases]))

    deterministic_attrs, responsive_ambiguous = html_resource_evidence(
        '<img src="required.svg">'
        '<img srcset="a.png 1x,b.png 2x">'
    )
    check("responsive ambiguity preserves unrelated deterministic resources",
          deterministic_attrs == {"required.svg"} and responsive_ambiguous,
          f"attrs={deterministic_attrs} ambiguous={responsive_ambiguous}")
    hidden_responsive_attrs, hidden_responsive_ambiguous = html_resource_evidence(
        '<div hidden><img srcset="a.png 1x,b.png 2x"></div>'
    )
    check("hidden responsive resources do not degrade asset evidence",
          hidden_responsive_attrs == set() and not hidden_responsive_ambiguous,
          f"attrs={hidden_responsive_attrs} "
          f"ambiguous={hidden_responsive_ambiguous}")
    svg_then_asset, svg_then_asset_ambiguous = html_resource_evidence(
        '<svg viewBox="0 0 100 100"><path d="M0 0L10 10" /></svg>'
        '<img src="required.svg">'
    )
    check("self-closing SVG graphics preserve following resource evidence",
          svg_then_asset == {"required.svg"} and not svg_then_asset_ambiguous,
          f"attrs={svg_then_asset} ambiguous={svg_then_asset_ambiguous}")
    check("coverage rejects omitted image assets",
          len(missing) == 1 and "missing-shot.png" in missing[0], str(missing))
    hidden, _, _ = coverage_issues(
        {"images": ["hidden-shot.png", "linked-only.png"]}, "", attrs
    )
    check("coverage ignores assets in templates and plain links",
          len(hidden) == 2, f"issues={hidden} attrs={attrs}")
    wrong_origin, _, _ = coverage_issues(
        ["https://brand.example/logo.svg"],
        "",
        html_resource_attributes('<img src="https://other.example/logo.svg">'),
        root_path="brief.required_assets",
        force_assets=True,
    )
    same_origin, _, _ = coverage_issues(
        ["https://brand.example/logo.svg?v=approved"],
        "",
        html_resource_attributes('<img src="https://brand.example/logo.svg?v=cache">'),
        root_path="brief.required_assets",
        force_assets=True,
    )
    check("required absolute assets cannot be impersonated by the same path on another host",
          len(wrong_origin) == 1, str(wrong_origin))
    check("required absolute assets tolerate cache-query changes on the same origin and path",
          same_origin == [], str(same_origin))

    local_absolute_remote_copy, _, _ = coverage_issues(
        ["/approved/brand/logo.svg"],
        "",
        html_resource_attributes(
            '<img src="https://attacker.example/approved/brand/logo.svg">'
        ),
        root_path="brief.required_assets",
        force_assets=True,
    )
    local_absolute_exact, _, _ = coverage_issues(
        ["/approved/brand/logo.svg"],
        "",
        html_resource_attributes('<img src="/approved/brand/logo.svg">'),
        root_path="brief.required_assets",
        force_assets=True,
    )
    check("required local absolute assets reject remote path impersonation",
          len(local_absolute_remote_copy) == 1 and local_absolute_exact == [],
          f"remote={local_absolute_remote_copy} exact={local_absolute_exact}")

    malformed_hidden_attrs = html_resource_attributes(
        '<div hidden><img></img><img src="hidden-after-void.svg"></div>'
        '<img hidden src="hidden-void.svg">'
        '<img src="visible.svg">'
    )
    check("resource parser keeps hidden scope across a void closing tag",
          malformed_hidden_attrs == {"visible.svg"},
          str(sorted(malformed_hidden_attrs)))

    self_closing_hidden_attrs = html_resource_attributes(
        '<div hidden/><img src="hidden-self-closing.svg"></div>'
        '<img src="visible-after-self-closing.svg">'
    )
    check("resource parser follows HTML semantics for non-void self-closing tags",
          self_closing_hidden_attrs == {"visible-after-self-closing.svg"},
          str(sorted(self_closing_hidden_attrs)))

    css_hidden_attrs = html_resource_attributes(
        '<style>.concealed img { display: none } '
        '.escaped { d\\69splay: n\\6f ne } '
        '[hidden] { display: none } '
        ':root { --fallback-display: none }</style>'
        '<div class="concealed"><img src="hidden-by-selector.svg"></div>'
        '<img style="display/**/: none" src="hidden-by-inline-comment.svg">'
        '<img class="escaped" src="hidden-by-css-escape.svg">'
        '<img hidden src="hidden-by-attribute.svg">'
        '<img src="visible-after-css.svg">'
    )
    check("resource parser fails closed on compound and comment-split hidden CSS",
          css_hidden_attrs == {"visible-after-css.svg"},
          str(sorted(css_hidden_attrs)))

    ambiguous_css_cases = [
        (
            '<style>div:not(.show) img { display: none }</style>'
            '<div><img src="hidden-by-not.svg"></div>',
            "hidden-by-not.svg",
        ),
        (
            '<style>[data-state^="hid"] { display: none }</style>'
            '<p data-state="hidden">HIDDEN-BY-ATTR-OPERATOR</p>',
            "HIDDEN-BY-ATTR-OPERATOR",
        ),
        (
            '<style>.marker + p { display: none }</style>'
            '<span class="marker"></span><p>HIDDEN-BY-SIBLING</p>',
            "HIDDEN-BY-SIBLING",
        ),
        (
            '<style>.hidden-by-var { --hide: none; display: var(--hide) }</style>'
            '<img class="hidden-by-var" src="hidden-by-var.svg">',
            "hidden-by-var.svg",
        ),
        (
            '<style>.h\\69 dden { visibility: collapse }</style>'
            '<img class="hidden" src="hidden-by-selector-escape.svg">',
            "hidden-by-selector-escape.svg",
        ),
        (
            '<link rel="style&#115;heet" '
            'href="data:text/css,.x%7Bdisplay%3Anone%7D">'
            '<img class="x" src="hidden-by-encoded-stylesheet.svg">',
            "hidden-by-encoded-stylesheet.svg",
        ),
        (
            '<style>@\\69mport url("data:text/css,.x%7Bdisplay%3Anone%7D")</style>'
            '<img class="x" src="hidden-by-escaped-import.svg">',
            "hidden-by-escaped-import.svg",
        ),
    ]
    ambiguous_results = [
        (
            marker,
            visible_html_text(raw, fail_closed=True),
            html_resource_attributes(raw),
        )
        for raw, marker in ambiguous_css_cases
    ]
    check("unsupported hiding CSS fails closed instead of partially matching",
          all(
              marker not in text and marker not in attrs
              for marker, text, attrs in ambiguous_results
          ),
          str(ambiguous_results))

    responsive_attrs = html_resource_attributes(
        '<picture>'
        '<source media="(min-width:99999px)" srcset="never-selected.svg">'
        '<img src="actual.svg" srcset="candidate-1.svg 1x, candidate-2.svg 2x">'
        '</picture>'
    )
    check("resource parser excludes unresolved responsive candidates",
          responsive_attrs == set(),
          str(sorted(responsive_attrs)))
    picture_fallback_attrs = html_resource_attributes(
        '<picture>'
        '<source media="(min-width:1px)" srcset="actual.svg">'
        '<img src="required-fallback.svg">'
        '</picture>'
    )
    bare_source_attrs = html_resource_attributes(
        '<source srcset="bare-unrendered.svg">'
    )
    check("picture fallbacks and bare sources cannot prove required assets",
          picture_fallback_attrs == set() and bare_source_attrs == set(),
          f"picture={sorted(picture_fallback_attrs)} bare={sorted(bare_source_attrs)}")

    remote_base, _, _ = coverage_issues(
        ["approved/logo.svg"],
        "",
        html_resource_attributes(
            '<base href="https://attacker.example/">'
            '<img src="approved/logo.svg">'
        ),
        root_path="brief.required_assets",
        force_assets=True,
    )
    check("required local assets reject a remote base URL",
          len(remote_base) == 1,
          str(remote_base))


def test_coverage_caps_adversarial_reports() -> None:
    from content import MAX_COVERAGE_ISSUES, MAX_COVERAGE_VALUES, coverage_issues

    missing, _, _ = coverage_issues({"values": list(range(1000))}, "")
    oversized, _, _ = coverage_issues({"values": ["present"] * (MAX_COVERAGE_VALUES + 1)}, "present")
    check("coverage caps missing-value reports",
          len(missing) == MAX_COVERAGE_ISSUES + 1
          and "issue limit" in missing[-1], f"issues={len(missing)}")
    check("coverage caps the number of atomic values",
          len(oversized) == 1 and "too many atomic values" in oversized[0], str(oversized[-2:]))


# --------------------------- MCP server ---------------------------

def test_mcp_server_stdio_protocol() -> None:
    """The server must speak newline-delimited JSON-RPC with nothing else on stdout."""
    script = REPO_ROOT / "scripts" / "mcp_server.py"
    msgs = [
        {"jsonrpc": "2.0", "method": "initialize", "params": {}},
        {"jsonrpc": "2.0", "id": 1, "method": "initialize",
         "params": {"protocolVersion": "2025-03-26"}},
        {"jsonrpc": "2.0", "method": "notifications/initialized"},
        {"jsonrpc": "2.0", "id": None, "method": "ping"},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
        {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
         "params": {"name": "kami_templates", "arguments": {}}},
        {"jsonrpc": "2.0", "id": 4, "method": "tools/call",
         "params": {"name": "kami_doctor", "arguments": {}}},
        {"jsonrpc": "2.0", "id": 5, "method": "tools/call",
         "params": {"name": "nope", "arguments": {}}},
    ]
    stdin = "".join(json.dumps(m) + "\n" for m in msgs)
    result = subprocess.run(
        [sys.executable, str(script)],
        input=stdin, capture_output=True, text=True, cwd=REPO_ROOT, timeout=60,
    )
    try:
        replies = {m.get("id"): m for m in map(json.loads, result.stdout.strip().splitlines())}
    except json.JSONDecodeError:
        check("mcp server stdout is newline-delimited JSON", False, result.stdout[:200])
        return
    init = replies.get(1, {}).get("result", {})
    check("mcp initialize echoes protocol version and names the server",
          init.get("protocolVersion") == "2025-03-26"
          and init.get("serverInfo", {}).get("name") == "kami",
          json.dumps(init)[:200])
    tools = [t["name"] for t in replies.get(2, {}).get("result", {}).get("tools", [])]
    check("mcp tools/list exposes the five kami tools",
          tools == ["kami_templates", "kami_doctor", "kami_render", "kami_check", "kami_screenshot"],
          str(tools))
    body = replies.get(3, {}).get("result", {}).get("content", [{}])[0].get("text", "{}")
    payload = json.loads(body)
    check("mcp kami_templates returns registries and schema types",
          set(payload.get("document_templates", {})) == set(HTML_TEMPLATES)
          and set(payload.get("pptx_templates", {})) == {"slides", "slides-en"}
          and set(payload.get("marp_templates", {})) == set(MARP_TEMPLATES)
          and payload.get("content_schema_types"),
          body[:200])
    doctor_body = replies.get(4, {}).get("result", {}).get("content", [{}])[0].get("text", "{}")
    doctor = json.loads(doctor_body)
    check("mcp kami_doctor reports dependencies, fonts, and capabilities",
          isinstance(doctor.get("ok"), bool)
          and len(doctor.get("dependencies", [])) >= 3
          and len(doctor.get("fonts", [])) >= 3
          and "pdf_visual_review" in doctor.get("capabilities", {})
          and "strict_math" in doctor.get("capabilities", {}),
          doctor_body[:300])
    check("mcp unknown tool returns a JSON-RPC error",
          "error" in replies.get(5, {}), json.dumps(replies.get(5, {}))[:200])
    check("mcp explicit null id remains a request",
          replies.get(None, {}).get("result") == {}, str(replies.get(None)))
    check("mcp notifications produced no reply", len(replies) == 6, str(sorted(
        ("null" if key is None else str(key)) for key in replies
    )))


def test_mcp_check_returns_stable_findings_and_coverage() -> None:
    from mcp_server import CHECK_REGISTRY, tool_check

    with tempfile.TemporaryDirectory() as d:
        clean = Path(d) / "clean.html"
        broken = Path(d) / "broken.html"
        math_broken = Path(d) / "math-broken.html"
        clean.write_text("<html><body><p>Ready</p></body></html>", encoding="utf-8")
        broken.write_text("<html><body><p>{{ missing }}</p></body></html>", encoding="utf-8")
        math_broken.write_text(
            r"<html><body><p>\(x^2\)</p></body></html>",
            encoding="utf-8",
        )
        invalid_content = Path(d) / "invalid-content.json"
        invalid_content.write_text(
            json.dumps({"type": "letter", "lang": "en", "content": {}}),
            encoding="utf-8",
        )
        valid_content = Path(d) / "valid-content.json"
        valid_content.write_text(json.dumps({
            "type": "letter",
            "lang": "en",
            "content": {
                "sender": "Ada Lovelace, London",
                "date": "2026-07-13",
                "recipient": "Charles Babbage",
                "salutation": "Dear Charles,",
                "paragraphs": [
                    "I write to state my purpose in one sentence: the engine deserves a program of its own.",
                    "The evidence sits in the notes: fifty operations, one loop, and a table the machine can follow.",
                    "My ask is specific: review the table this month so we can test it on the mill.",
                ],
                "signoff": "Sincerely,",
                "signature": "Ada",
            },
        }), encoding="utf-8")
        clean_result = tool_check({"path": str(clean)})
        broken_result = tool_check({"path": str(broken)})
        math_broken_result = tool_check({"path": str(math_broken)})
        invalid_content_result = tool_check({
            "path": str(clean), "content": str(invalid_content),
        })
        missing_coverage_result = tool_check({
            "path": str(clean), "content": str(valid_content),
        })

    check("MCP check registry carries unique stable rule IDs",
          CHECK_REGISTRY and clean_result["ruleset_version"] == 3
          and len(CHECK_REGISTRY) == len(set(CHECK_REGISTRY))
          and all({"scope", "severity", "required_engine", "explanation"} <= set(rule)
                  for rule in CHECK_REGISTRY.values()),
          str(CHECK_REGISTRY))
    check("MCP clean check returns coverage without findings",
          clean_result["ok"] is True
          and clean_result["degraded"] is False
          and clean_result["findings"] == []
          and [item["id"] for item in clean_result["coverage"]]
          == ["html.placeholders", "html.math", "html.markdown-residue"]
          and clean_result["report"],
          json.dumps(clean_result)[:500])
    check("MCP failed check returns a stable finding and legacy report",
          broken_result["ok"] is False
          and broken_result["findings"][0]["id"] == "html.placeholders"
          and broken_result["findings"][0]["status"] == "failed"
          and "placeholder" in broken_result["report"].lower(),
          json.dumps(broken_result)[:500])
    check("MCP HTML check rejects unrendered standard LaTeX",
          math_broken_result["ok"] is False
          and [finding["id"] for finding in math_broken_result["findings"]]
          == ["html.math"],
          json.dumps(math_broken_result)[:500])
    invalid_ids = [item["id"] for item in invalid_content_result["findings"]]
    invalid_coverage = {
        item["id"]: item for item in invalid_content_result["coverage"]
    }
    check("MCP labels invalid content IR as contract failure before coverage",
          "content.contract" in invalid_ids
          and "content.coverage" not in invalid_ids
          and invalid_coverage["content.coverage"]["status"] == "not_run"
          and invalid_coverage["content.coverage"]["blocked_by"]
          == "content.contract",
          json.dumps(invalid_content_result)[:800])
    missing_ids = [item["id"] for item in missing_coverage_result["findings"]]
    missing_statuses = {
        item["id"]: item["status"]
        for item in missing_coverage_result["coverage"]
    }
    check("MCP runs coverage once after a valid content contract",
          missing_ids == ["content.coverage"]
          and missing_statuses["content.contract"] == "passed"
          and missing_statuses["content.coverage"] == "failed",
          json.dumps(missing_coverage_result)[:800])


def test_mcp_server_rejects_bad_frames_without_exiting() -> None:
    """Wrong-shaped JSON-RPC frames return errors and do not kill the server."""
    script = REPO_ROOT / "scripts" / "mcp_server.py"
    msgs = [
        [],
        {"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": "bad"},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/call",
         "params": {"name": "kami_templates", "arguments": ["bad"]}},
        {"jsonrpc": "2.0", "id": 3, "method": "ping"},
    ]
    result = subprocess.run(
        [sys.executable, str(script)],
        input="".join(json.dumps(m) + "\n" for m in msgs),
        capture_output=True, text=True, cwd=REPO_ROOT, timeout=60,
    )
    replies = [json.loads(line) for line in result.stdout.strip().splitlines()]
    by_id = {reply.get("id"): reply for reply in replies}
    check("mcp malformed frames keep the process alive",
          result.returncode == 0 and len(replies) == 4 and "result" in by_id.get(3, {}),
          (result.stdout + result.stderr)[:400])
    check("mcp wrong params and arguments return invalid-params errors",
          by_id.get(1, {}).get("error", {}).get("code") == -32602
          and by_id.get(2, {}).get("error", {}).get("code") == -32602,
          result.stdout[:400])


def test_mcp_all_tools_succeed_over_stdio() -> None:
    """Exercise render, check, and screenshot through the installed protocol path."""
    try:
        from optional_deps import require_pypdf_reader, require_weasyprint_html
        require_weasyprint_html()
        require_pypdf_reader()
        require_pymupdf()
    except MissingDepError as exc:
        skip("MCP all-tools stdio success path", str(exc), ci_required=True)
        return

    script = REPO_ROOT / "scripts" / "mcp_server.py"
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        html = root / "source.html"
        pdf = root / "output.pdf"
        html.write_text(
            "<!doctype html><html><head><style>"
            "@page{size:A4;margin:20mm}body{font-family:serif}"
            "</style></head><body><h1>Kami MCP smoke</h1><p>Rendered.</p></body></html>",
            encoding="utf-8",
        )
        msgs = [
            {"jsonrpc": "2.0", "id": 1, "method": "initialize",
             "params": {"protocolVersion": "2025-06-18"}},
            {"jsonrpc": "2.0", "id": 2, "method": "tools/call",
             "params": {"name": "kami_render", "arguments": {
                 "html": str(html), "out": str(pdf)}}},
            {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
             "params": {"name": "kami_check", "arguments": {"path": str(html)}}},
            {"jsonrpc": "2.0", "id": 4, "method": "tools/call",
             "params": {"name": "kami_screenshot", "arguments": {"pdf": str(pdf)}}},
        ]
        result = subprocess.run(
            [sys.executable, str(script)],
            input="".join(json.dumps(m) + "\n" for m in msgs),
            capture_output=True, text=True, cwd=REPO_ROOT, timeout=120,
        )
        try:
            replies = {m.get("id"): m for m in map(json.loads, result.stdout.strip().splitlines())}
            payloads = {
                reply_id: json.loads(
                    replies[reply_id]["result"]["content"][0]["text"]
                )
                for reply_id in (2, 3, 4)
            }
        except (KeyError, json.JSONDecodeError) as exc:
            check("MCP all-tools success path returns JSON results", False,
                  f"{exc}: {(result.stdout + result.stderr)[:500]}")
            return

        render_result = payloads[2]
        check_result = payloads[3]
        screenshot_result = payloads[4]
        check("MCP render succeeds over stdio",
              result.returncode == 0 and render_result.get("pages") == 1 and pdf.is_file(),
              json.dumps(render_result)[:300])
        check("MCP check succeeds over stdio",
              check_result.get("ok") is True and check_result.get("exit_code") == 0,
              json.dumps(check_result)[:300])
        page_paths = [Path(path) for path in screenshot_result.get("pages", [])]
        check("MCP screenshot returns evidence without a false perceptual verdict",
              "ok" not in screenshot_result
              and screenshot_result.get("rasterized") is True
              and screenshot_result.get("review_pending") is True
              and screenshot_result.get("font_check", {}).get("ok") is True
              and len(page_paths) == 1 and all(path.is_file() for path in page_paths),
              json.dumps(screenshot_result)[:500])


def test_mcp_render_guards_source_and_output_types() -> None:
    from mcp_server import tool_render

    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        html = root / "source.html"
        html.write_text("<html><body>safe source</body></html>", encoding="utf-8")
        original = html.read_bytes()
        hardlink = root / "hardlink.pdf"
        hardlink.hardlink_to(html)
        victim = root / "victim.txt"
        victim.write_bytes(b"do-not-overwrite")
        symlink = root / "symlink.pdf"
        symlink.symlink_to(victim)
        rejected = 0
        for out in (html, hardlink, symlink, root / "not-pdf.txt"):
            try:
                tool_render({"html": str(html), "out": str(out)})
            except ValueError:
                rejected += 1
        check("mcp render rejects source aliases and non-PDF outputs",
              rejected == 4 and html.read_bytes() == original
              and victim.read_bytes() == b"do-not-overwrite",
              f"rejected={rejected} source_changed={html.read_bytes() != original}")


def test_visual_rejects_empty_pdf_and_bad_dpi() -> None:
    try:
        from pypdf import PdfWriter
        from visual import render_pages
    except ImportError:
        skip("visual empty-PDF guard", "pypdf unavailable", ci_required=True)
        return

    with tempfile.TemporaryDirectory() as d:
        pdf = Path(d) / "empty.pdf"
        evidence = Path(d) / "empty-visual"
        evidence.mkdir()
        old_page = evidence / "page-01.png"
        old_page.write_bytes(b"last-good-run")
        writer = PdfWriter()
        writer.write(str(pdf))
        errors = 0
        for dpi in (1, -1, 301):
            try:
                render_pages(pdf, dpi=dpi)
            except ValueError:
                errors += 1
        try:
            render_pages(pdf, dpi=110)
        except ValueError as exc:
            empty_error = "no pages" in str(exc)
        else:
            empty_error = False
        check("visual rejects empty PDFs and out-of-range DPI",
              empty_error and errors == 3,
              f"empty_error={empty_error} dpi_errors={errors}")
        check("failed visual render preserves last good evidence",
              old_page.read_bytes() == b"last-good-run")

        symlink_target = Path(d) / "elsewhere"
        symlink_target.mkdir()
        symlink_output = Path(d) / "linked-visual"
        symlink_output.symlink_to(symlink_target, target_is_directory=True)
        try:
            render_pages(pdf, out_dir=symlink_output, dpi=110)
        except ValueError as exc:
            symlink_rejected = "symbolic link" in str(exc)
        else:
            symlink_rejected = False
        check("visual rejects a symbolic-link evidence directory", symlink_rejected)

        huge_pdf = Path(d) / "huge.pdf"
        huge_writer = PdfWriter()
        huge_writer.add_blank_page(width=5000, height=5000)
        huge_writer.write(str(huge_pdf))
        try:
            render_pages(huge_pdf, dpi=300)
        except ValueError as exc:
            huge_rejected = "pixels" in str(exc)
        else:
            huge_rejected = False
        check("visual rejects an oversized raster page before rendering", huge_rejected)


def test_document_checks_reject_empty_pdf() -> None:
    """A readable PDF with zero pages is not a successfully checked artifact."""
    try:
        from pypdf import PdfWriter
        from checks import check_density, check_orphans
        from mcp_server import tool_check
    except ImportError:
        skip("empty-PDF document checks", "render dependencies unavailable", ci_required=True)
        return

    with tempfile.TemporaryDirectory() as d:
        pdf = Path(d) / "empty.pdf"
        PdfWriter().write(str(pdf))
        results = [
            silently(check_markdown_residue, [str(pdf)]),
            silently(check_orphans, [str(pdf)]),
            silently(check_density, [str(pdf)]),
        ]
        mcp_result = tool_check({"path": str(pdf)})
        check("PDF checks reject zero-page artifacts",
              results == [2, 2, 2], str(results))
        check("MCP check rejects zero-page artifacts",
              not mcp_result["ok"]
              and all(rule["status"] == "degraded" for rule in mcp_result["coverage"]),
              str(mcp_result))


def _test_functions():
    tests = []
    for name, func in globals().items():
        if not name.startswith("test_") or not callable(func):
            continue
        if getattr(func, "__module__", None) != __name__:
            continue
        code = getattr(func, "__code__", None)
        if code is None:
            continue
        tests.append((code.co_firstlineno, name, func))
    return [(name, func) for _, name, func in sorted(tests)]


def main() -> int:
    for name, func in _test_functions():
        signature = inspect.signature(func)
        if signature.parameters:
            params = ", ".join(signature.parameters)
            check(f"{name} has no parameters", False, f"parameters: {params}")
            continue
        func()
    print()
    print(f"Passed: {_PASS} | Skipped: {_SKIP} | Failed: {_FAIL}")
    return 0 if _FAIL == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
