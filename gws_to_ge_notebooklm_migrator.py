#!/usr/bin/env python3
"""GWS NotebookLM -> SharePoint (.docx/.md) -> Gemini Enterprise (NotebookLM Enterprise)

3-Fold Technical Admin Migration Tool & Web Workbench (Single-User & Multi-User
Batch Engine).

Fold 1: Extract GWS Takeout ZIP(s), Multi-ZIP Wave Bundles, or Customer Takeout
GCS Buckets
        (`gs://bucket/...`) -> Parse per-user Notebooks, Sources (HTML +
        metadata.json),
        Original Website/YouTube URLs, Notes, and Studio Artifacts.
Fold 2: Convert & Package Sources into a Multi-User SharePoint/OneDrive-ready
`.docx` hierarchy
        (`SharePoint_Ready_Notebooks.zip`) with strict UTF-8 ZIP flags (`0x800`)
        + Microsoft
        SharePoint Migration Tool (`sharepoint_spmt_manifest.csv`) + PnP
        PowerShell & Microsoft
        Graph API bulk uploader scripts.
Fold 3: Bulk-create Notebooks across 1 to 100+ users on Gemini Enterprise
        (`discoveryengine.googleapis.com/v1alpha`) using a bounded
        ThreadPoolExecutor,
        HTTP 429/5xx exponential backoff, idempotent `batch_checkpoint.json`
        resume ledger,
        and automatic per-user `notebooks:share`.
"""

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv
import email.parser
import html as html_lib
from http.server import BaseHTTPRequestHandler, HTTPServer
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import threading
import time
from urllib.parse import urlparse, urlunparse
import zipfile

from bs4 import BeautifulSoup
import requests

WORKSPACE_DIR = Path(
    os.environ.get("NBLM_WORKSPACE_DIR", "/tmp/nblm_migration_workspace")
)
WORKSPACE_DIR.mkdir(parents=True, exist_ok=True)
CHECKPOINT_LOCK = threading.Lock()


def sanitize_filename(name: str) -> str:
  """Sanitizes a string to be safe for SharePoint, OneDrive, Windows, and macOS paths."""
  cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip()
  cleaned = re.sub(r"\s+", " ", cleaned)
  cleaned = cleaned.rstrip(". ")
  return cleaned[:180] or "Untitled"


def write_utf8_zip_bytes(
    zf: zipfile.ZipFile, arcname: str, data: bytes
) -> None:
  """Writes a file to a ZipFile with explicit UTF-8 flag_bits (0x800) so Windows/macOS never create duplicate CP437 folders."""
  zinfo = zipfile.ZipInfo(filename=arcname)
  zinfo.compress_type = zipfile.ZIP_DEFLATED
  zinfo.flag_bits |= 0x0800  # Bit 11 = UTF-8 filename encoding
  zinfo.external_attr = 0o644 << 16
  zf.writestr(zinfo, data)


def extract_original_url_from_html(title: str, html_content: str) -> str:
  """Infers the original scraped page URL from preserved <a href> links in NotebookLM's Takeout HTML."""
  if not html_content:
    return ""
  soup = BeautifulSoup(html_content, "html.parser")
  hrefs = []
  for a in soup.find_all("a"):
    h = (a.get("href") or "").strip()
    if h.startswith("http://") or h.startswith("https://"):
      hrefs.append(h)
  if not hrefs:
    return ""

  clean_title = re.sub(r"[^a-zA-Z0-9\s]", " ", title.lower())
  stop_words = {
      "the",
      "and",
      "for",
      "with",
      "from",
      "how",
      "what",
      "why",
      "com",
      "org",
      "net",
      "wikipedia",
  }
  title_tokens = [
      w for w in clean_title.split() if len(w) >= 3 and w not in stop_words
  ]

  parts = re.split(r"\s+[-|]\s+", title)
  site_hint = parts[-1].lower().strip() if len(parts) > 1 else ""
  site_tokens = [w for w in re.sub(r"[^a-z0-9]", " ", site_hint).split() if w]

  best_url = ""
  best_score = -1.0

  for idx, raw_url in enumerate(hrefs):
    try:
      parsed = urlparse(raw_url)
    except Exception:
      continue
    host = (parsed.netloc or "").lower()
    path = (parsed.path or "").lower()
    fragment = (parsed.fragment or "").lower()

    if any(
        bad in host
        for bad in [
            "onelink.me",
            "sng.link",
            "twitter.com",
            "facebook.com",
            "linkedin.com",
            "t.co",
            "accounts.google",
        ]
    ):
      continue
    if any(
        bad in path
        for bad in [
            "/login",
            "/signup",
            "/cart",
            "/customer_authentication",
            "/special:",
            "/wikipedia:",
            "/portal:",
            "/main_page",
            "/help:",
            "/talk:",
        ]
    ):
      continue
    if parsed.query and (
        "action=edit" in parsed.query or "printable=yes" in parsed.query
    ):
      continue

    score = 0.0
    if fragment in (
        "bodycontent",
        "maincontent",
        "content",
        "main",
        "overview",
        "course-outline",
        "top",
    ):
      score += 50.0
      if idx == 0:
        score += 30.0
    elif fragment:
      score += 10.0

    path_clean = re.sub(r"[^a-z0-9]", " ", path)
    matched_tokens = sum(1 for tok in title_tokens if tok in path_clean)
    if title_tokens:
      score += (matched_tokens / len(title_tokens)) * 40.0
    score += matched_tokens * 8.0

    if "wikipedia.org/wiki/" in (host + path):
      wiki_slug = (
          title.split(" - Wikipedia")[0].strip().replace(" ", "_").lower()
      )
      if path == f"/wiki/{wiki_slug}":
        score += 100.0

    if site_tokens and any(st in host for st in site_tokens):
      score += 15.0

    if path in ("", "/", "/en", "/courses", "/blog"):
      score -= 20.0

    clean_query = ""
    if (
        parsed.query
        and "utm_" not in parsed.query
        and "locale=" not in parsed.query
    ):
      clean_query = parsed.query
    canonical = urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path.rstrip("/"),
        "",
        clean_query,
        "",
    ))

    if score > best_score:
      best_score = score
      best_url = canonical

  return best_url


def html_to_blocks(html_content: str) -> list[dict]:
  """Parses NotebookLM exported HTML into structured blocks."""
  soup = BeautifulSoup(html_content or "", "html.parser")
  blocks = []
  for el in soup.find_all(
      ["h1", "h2", "h3", "h4", "p", "li", "pre", "blockquote"]
  ):
    text = el.get_text(" ", strip=True)
    if not text:
      continue
    tag = el.name.lower()
    if tag in ("h1", "h2", "h3", "h4"):
      blocks.append({"type": "heading", "level": int(tag[1]), "text": text})
    elif tag == "li":
      blocks.append({"type": "bullet", "text": text})
    else:
      blocks.append({"type": "paragraph", "text": text})

  if not blocks:
    raw_text = soup.get_text("\n", strip=True)
    for line in raw_text.splitlines():
      line = line.strip()
      if line:
        blocks.append({"type": "paragraph", "text": line})
  return blocks


def build_minimal_docx_bytes(title: str, meta: dict, blocks: list[dict]) -> bytes:
  """Generates a Microsoft Word (.docx) OOXML package in memory without external dependencies."""
  content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

  rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

  def make_para(
      text: str, bold: bool = False, size_half_pts: int = 22, color: str = "202124"
  ) -> str:
    safe_text = html_lib.escape(
        re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    )
    rpr = f'<w:rPr><w:sz w:val="{size_half_pts}"/><w:color w:val="{color}"/>'
    if bold:
      rpr += "<w:b/>"
    rpr += "</w:rPr>"
    return (
        f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{safe_text}</w:t></w:r></w:p>'
    )

  body_paras = [
      make_para(title, bold=True, size_half_pts=36, color="1A73E8"),
      make_para(
          f"Original Source Type: {meta.get('source_type', 'DOCUMENT')} | Owner:"
          f" {meta.get('owner_email', 'N/A')} | Added:"
          f" {meta.get('added_timestamp', 'N/A')}",
          bold=False,
          size_half_pts=18,
          color="5F6368",
      ),
  ]
  if meta.get("source_url"):
    body_paras.append(
        make_para(
            f"Original Source URL: {meta['source_url']}",
            bold=True,
            size_half_pts=20,
            color="1A73E8",
        )
    )
  if meta.get("youtube_channel"):
    body_paras.append(
        make_para(
            f"YouTube Channel: {meta['youtube_channel']}",
            bold=False,
            size_half_pts=18,
            color="5F6368",
        )
    )

  for b in blocks:
    if b["type"] == "heading":
      sz = {1: 32, 2: 28, 3: 24}.get(b.get("level", 2), 24)
      body_paras.append(
          make_para(b["text"], bold=True, size_half_pts=sz, color="174EA6")
      )
    elif b["type"] == "bullet":
      body_paras.append(make_para(f"• {b['text']}", bold=False, size_half_pts=22))
    else:
      body_paras.append(make_para(b["text"], bold=False, size_half_pts=22))

  document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {"".join(body_paras)}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""

  buf = io.BytesIO()
  with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
    write_utf8_zip_bytes(
        zf, "[Content_Types].xml", content_types_xml.encode("utf-8")
    )
    write_utf8_zip_bytes(zf, "_rels/.rels", rels_xml.encode("utf-8"))
    write_utf8_zip_bytes(zf, "word/document.xml", document_xml.encode("utf-8"))
  return buf.getvalue()


def infer_user_email_from_path(
    rel_path: str, default_email: str = "admin@andresvilla.altostrat.com"
) -> str:
  """Detects GWS user email from Customer Takeout GCS path (`<prefix>/<user@domain>/Takeout/NotebookLM/...`)

  or filename (`takeout-<user@domain>.zip`).
  """
  email_match = re.search(
      r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", rel_path
  )
  if email_match:
    return email_match.group(1).lower()
  return default_email


def parse_user_mapping_csv_text(csv_text: str) -> dict[str, dict]:
  """Parses a CSV string mapping gws_email -> m365_upn, ge_target_email, sharepoint_site_url."""
  mapping = {}
  if not csv_text or not csv_text.strip():
    return mapping
  reader = csv.DictReader(io.StringIO(csv_text.strip()))
  for row in reader:
    norm_row = {
        (k or "").strip().lower(): (v or "").strip() for k, v in row.items()
    }
    gws_email = (
        norm_row.get("gws_email")
        or norm_row.get("gws_source_email")
        or norm_row.get("source_email")
        or norm_row.get("email")
        or ""
    ).lower()
    if not gws_email:
      continue
    m365_upn = (
        norm_row.get("m365_upn")
        or norm_row.get("m365_sharepoint_upn")
        or norm_row.get("target_upn")
        or gws_email
    )
    ge_email = (
        norm_row.get("ge_target_email")
        or norm_row.get("gemini_enterprise_target_email")
        or m365_upn
    )
    sp_url = (
        norm_row.get("sharepoint_site_url")
        or norm_row.get("sharepoint_onedrive_destination_url")
        or f"https://contoso.sharepoint.com/sites/GeminiNotebooks/Shared Documents/{m365_upn}"
    )
    mapping[gws_email] = {
        "gws_email": gws_email,
        "m365_upn": m365_upn,
        "ge_target_email": ge_email,
        "sharepoint_site_url": sp_url,
    }
  return mapping


PNP_POWERSHELL_SCRIPT = """# PnP PowerShell Bulk SharePoint / OneDrive Uploader for Migrated NotebookLM (.docx)
# Reads sharepoint_spmt_manifest.csv and uploads each user's converted .docx sources to SharePoint Online / OneDrive.
param(
    [Parameter(Mandatory=$false)][string]$ManifestCsv = "./sharepoint_spmt_manifest.csv",
    [Parameter(Mandatory=$false)][string]$ClientId = $env:PNP_CLIENT_ID
)

Import-Module PnP.PowerShell -ErrorAction Stop
$rows = Import-Csv -Path $ManifestCsv

foreach ($row in $rows) {
    Write-Host "Uploading Notebook folder $($row.SourcePath) -> $($row.TargetWebUrl)/$($row.TargetDocumentLibrary)/$($row.TargetSubFolder)" -ForegroundColor Cyan
    Connect-PnPOnline -Url $row.TargetWebUrl -Interactive -ClientId $ClientId
    $localFiles = Get-ChildItem -Path $row.SourcePath -Recurse -File
    foreach ($file in $localFiles) {
        $relDir = $file.DirectoryName.Substring($row.SourcePath.Length).TrimStart('\\', '/')
        $targetFolder = "$($row.TargetDocumentLibrary)/$($row.TargetSubFolder)/$relDir".TrimEnd('/')
        Add-PnPFile -Path $file.FullName -Folder $targetFolder -Values @{Title=$file.BaseName} | Out-Null
    }
}
Write-Host "All SharePoint / OneDrive NotebookLM folders uploaded successfully!" -ForegroundColor Green
"""


def build_manifest_from_extracted_dir(
    extract_dir: Path,
    output_dir: Path,
    default_owner_email: str = "admin@andresvilla.altostrat.com",
    user_mapping: dict[str, dict] | None = None,
) -> dict:
  """Scans `extract_dir` (which can hold 1 user's Takeout or 100 users' Takeouts from Multi-ZIP or GCS),

  applies `user_mapping` (if provided), generates SharePoint `.docx` files per
  user, creates `sharepoint_spmt_manifest.csv` + `upload_to_sharepoint_pnp.ps1`,
  and writes `SharePoint_Ready_Notebooks.zip` + `migration_manifest.json`.
  """
  sharepoint_dir = output_dir / "sharepoint_staging"
  if sharepoint_dir.exists():
    shutil.rmtree(sharepoint_dir)
  sharepoint_dir.mkdir(parents=True, exist_ok=True)
  user_mapping = user_mapping or {}

  notebooks_by_key = {}
  discovered_users = set()

  for root, dirs, files in os.walk(extract_dir):
    root_path = Path(root)
    if root_path.name in ("Sources", "Notes", "Artifacts", "Discovered Sources"):
      continue
    for f in sorted(files):
      if f.endswith(" metadata.json"):
        meta_path = root_path / f
        try:
          nb_meta = json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
          continue
        if "title" not in nb_meta and "emoji" not in nb_meta:
          continue

        rel_from_extract = root_path.relative_to(extract_dir).as_posix()
        owner_email = infer_user_email_from_path(
            rel_from_extract, default_owner_email
        )
        discovered_users.add(owner_email)

        u_map = user_mapping.get(owner_email, {})
        m365_upn = u_map.get("m365_upn") or owner_email
        ge_target_email = u_map.get("ge_target_email") or m365_upn
        sp_site_url = (
            u_map.get("sharepoint_site_url")
            or f"https://contoso.sharepoint.com/sites/GeminiNotebooks/Shared Documents/{m365_upn}"
        )

        nb_title = (
            nb_meta.get("title")
            or f[: -len(" metadata.json")]
            or "Untitled Notebook"
        ).strip()
        nb_emoji = nb_meta.get("emoji", "📓")
        nb_folder_name = sanitize_filename(nb_title)
        nb_key = f"{owner_email}::{nb_folder_name.lower()}"

        sp_nb_dir = sharepoint_dir / owner_email / nb_folder_name
        sp_sources_dir = sp_nb_dir / "Sources"
        sp_sources_dir.mkdir(parents=True, exist_ok=True)

        sources_list = []
        sources_src_dir = root_path / "Sources"
        if sources_src_dir.exists():
          for sf in sorted(os.listdir(sources_src_dir)):
            if sf.endswith(" metadata.json"):
              s_base = sf[: -len(" metadata.json")]
              s_meta_file = sources_src_dir / sf
              s_html_file = sources_src_dir / f"{s_base}.html"
              try:
                s_meta_raw = json.loads(s_meta_file.read_text(encoding="utf-8"))
              except Exception:
                s_meta_raw = {}
              inner_meta = s_meta_raw.get("metadata", {})
              source_type = inner_meta.get(
                  "originalSourceContentType", "SOURCE_CONTENT_TYPE_UNKNOWN"
              )
              added_ts = inner_meta.get("sourceAddedTimestamp", "")
              yt_meta = inner_meta.get("youtubeMetadata", {})
              video_id = yt_meta.get("videoId", "")
              yt_channel = yt_meta.get("channelName", "")
              youtube_url = (
                  f"https://www.youtube.com/watch?v={video_id}"
                  if video_id
                  else ""
              )

              html_str = (
                  s_html_file.read_text(encoding="utf-8", errors="replace")
                  if s_html_file.exists()
                  else ""
              )

              source_url = ""
              if youtube_url:
                source_url = youtube_url
              elif source_type == "SOURCE_CONTENT_TYPE_URL":
                source_url = extract_original_url_from_html(s_base, html_str)

              blocks = html_to_blocks(html_str)
              clean_s_name = sanitize_filename(s_base)
              meta_info = {
                  "source_type": source_type,
                  "owner_email": owner_email,
                  "added_timestamp": added_ts,
                  "source_url": source_url,
                  "youtube_url": youtube_url,
                  "youtube_channel": yt_channel,
              }

              docx_path = sp_sources_dir / f"{clean_s_name}.docx"
              docx_bytes = build_minimal_docx_bytes(s_base, meta_info, blocks)
              docx_path.write_bytes(docx_bytes)

              sources_list.append({
                  "title": s_base,
                  "source_type": source_type,
                  "added_timestamp": added_ts,
                  "source_url": source_url,
                  "youtube_url": youtube_url,
                  "youtube_channel": yt_channel,
                  "word_count": sum(len(b["text"].split()) for b in blocks),
                  "docx_path": str(docx_path),
                  "relative_sharepoint_path": (
                      f"{owner_email}/{nb_folder_name}/Sources/{clean_s_name}.docx"
                  ),
              })

        notes_list = []
        notes_src_dir = root_path / "Notes"
        if notes_src_dir.exists():
          for nf in sorted(os.listdir(notes_src_dir)):
            if nf.endswith(".html"):
              n_base = nf[: -len(".html")]
              n_html_file = notes_src_dir / nf
              html_str = n_html_file.read_text(encoding="utf-8", errors="replace")
              blocks = html_to_blocks(html_str)
              plain_text = "\n\n".join(b["text"] for b in blocks)
              clean_n_name = sanitize_filename(n_base)

              sp_notes_dir = sp_nb_dir / "Notes"
              sp_notes_dir.mkdir(parents=True, exist_ok=True)
              docx_path = sp_notes_dir / f"{clean_n_name}.docx"
              docx_path.write_bytes(
                  build_minimal_docx_bytes(
                      n_base,
                      {"source_type": "NOTE", "owner_email": owner_email},
                      blocks,
                  )
              )

              notes_list.append({
                  "title": n_base,
                  "content_text": plain_text,
                  "docx_path": str(docx_path),
                  "relative_sharepoint_path": (
                      f"{owner_email}/{nb_folder_name}/Notes/{clean_n_name}.docx"
                  ),
              })

        artifacts_list = []
        artifacts_src_dir = root_path / "Artifacts"
        if artifacts_src_dir.exists():
          for af in sorted(os.listdir(artifacts_src_dir)):
            if af.endswith(" metadata.json"):
              a_base = af[: -len(" metadata.json")]
              try:
                a_meta = json.loads(
                    (artifacts_src_dir / af).read_text(encoding="utf-8")
                )
              except Exception:
                a_meta = {}
              a_type = (
                  a_meta.get("artifactType")
                  or a_meta.get("appType")
                  or "STUDIO_ARTIFACT"
              )
              artifacts_list.append({
                  "title": a_meta.get("title") or a_base,
                  "artifact_type": a_type,
              })

        index_lines = [
            f"# {nb_emoji} {nb_title}",
            "",
            f"- **GWS Owner**: `{owner_email}`",
            f"- **Target M365 / SharePoint Owner**: `{m365_upn}`",
            f"- **Target Gemini Enterprise User**: `{ge_target_email}`",
            f"- **Total Sources**: {len(sources_list)}",
            "",
            "## Sources Inventory & Original URLs",
            "",
            "| # | Source Title | Original Type | Original URL / Origin | SharePoint File |",
            "|---|---|---|---|---|",
        ]
        for idx, s in enumerate(sources_list, 1):
          url_cell = (
              f"[{s['source_url']}]({s['source_url']})"
              if s.get("source_url")
              else "_Uploaded File_"
          )
          index_lines.append(
              f"| {idx} | {s['title']} | `{s['source_type']}` | {url_cell} |"
              f" `Sources/{Path(s['docx_path']).name}` |"
          )
        (sp_nb_dir / "Notebook_Sources_Index.md").write_text(
            "\n".join(index_lines) + "\n", encoding="utf-8"
        )

        if nb_key in notebooks_by_key:
          existing = notebooks_by_key[nb_key]
          existing_titles = {s["title"] for s in existing["sources"]}
          for s in sources_list:
            if s["title"] not in existing_titles:
              existing["sources"].append(s)
        else:
          notebooks_by_key[nb_key] = {
              "owner_email": owner_email,
              "m365_upn": m365_upn,
              "ge_target_email": ge_target_email,
              "sharepoint_site_url": sp_site_url,
              "title": nb_title,
              "emoji": nb_emoji,
              "metadata": nb_meta,
              "sharepoint_folder": f"{owner_email}/{nb_folder_name}",
              "sources": sources_list,
              "notes": notes_list,
              "artifacts": artifacts_list,
          }

  notebooks = list(notebooks_by_key.values())

  # If user_mapping has multiple users (e.g. a 10-100 user batch wave CSV) and only 1 template user was extracted,
  # or if multiple users exist in user_mapping, expand/apply the mapping rows so every mapped user gets their staged folder & SPMT entry!
  if len(user_mapping) > 1 and len(discovered_users) == 1:
    template_user = next(iter(discovered_users))
    template_nbs = [
        nb for nb in notebooks if nb["owner_email"] == template_user
    ]
    expanded_notebooks = []
    discovered_users = set()
    for g_email, u_info in user_mapping.items():
      discovered_users.add(g_email)
      for t_nb in template_nbs:
        nb_folder_name = sanitize_filename(t_nb["title"])
        user_nb_dir = sharepoint_dir / g_email / nb_folder_name
        if not user_nb_dir.exists():
          shutil.copytree(
              sharepoint_dir / template_user / nb_folder_name,
              user_nb_dir,
              dirs_exist_ok=True,
          )
        new_sources = []
        for s in t_nb["sources"]:
          s_copy = dict(s)
          s_copy["docx_path"] = str(
              user_nb_dir / "Sources" / Path(s["docx_path"]).name
          )
          s_copy["relative_sharepoint_path"] = (
              f"{g_email}/{nb_folder_name}/Sources/{Path(s['docx_path']).name}"
          )
          new_sources.append(s_copy)
        expanded_notebooks.append({
            "owner_email": g_email,
            "m365_upn": u_info["m365_upn"],
            "ge_target_email": u_info["ge_target_email"],
            "sharepoint_site_url": u_info["sharepoint_site_url"],
            "title": t_nb["title"],
            "emoji": t_nb["emoji"],
            "metadata": t_nb["metadata"],
            "sharepoint_folder": f"{g_email}/{nb_folder_name}",
            "sources": new_sources,
            "notes": t_nb["notes"],
            "artifacts": t_nb["artifacts"],
        })
    notebooks = expanded_notebooks

  # Write Microsoft SharePoint Migration Tool (SPMT) CSV manifest + PnP PowerShell script
  spmt_csv_path = sharepoint_dir / "sharepoint_spmt_manifest.csv"
  with open(spmt_csv_path, "w", newline="", encoding="utf-8") as cf:
    writer = csv.writer(cf)
    writer.writerow([
        "SourcePath",
        "TargetDocumentLibrary",
        "TargetSubFolder",
        "TargetWebUrl",
        "GWS_Owner_Email",
        "M365_SharePoint_UPN",
        "GE_Target_Email",
    ])
    for nb in notebooks:
      local_nb_folder = str(sharepoint_dir / nb["sharepoint_folder"])
      writer.writerow([
          local_nb_folder,
          "Documents",
          f"NotebookLM_Migrated/{sanitize_filename(nb['title'])}",
          nb["sharepoint_site_url"],
          nb["owner_email"],
          nb["m365_upn"],
          nb["ge_target_email"],
      ])

  pnp_script_path = sharepoint_dir / "upload_to_sharepoint_pnp.ps1"
  pnp_script_path.write_text(PNP_POWERSHELL_SCRIPT, encoding="utf-8")

  sp_zip_path = output_dir / "SharePoint_Ready_Notebooks.zip"
  single_user_only = len(discovered_users) <= 1
  with zipfile.ZipFile(sp_zip_path, "w", zipfile.ZIP_DEFLATED) as sp_zf:
    write_utf8_zip_bytes(
        sp_zf, "sharepoint_spmt_manifest.csv", spmt_csv_path.read_bytes()
    )
    write_utf8_zip_bytes(
        sp_zf, "upload_to_sharepoint_pnp.ps1", pnp_script_path.read_bytes()
    )
    for r, d, f_list in os.walk(sharepoint_dir):
      for f_item in sorted(f_list):
        full_p = Path(r) / f_item
        if full_p in (spmt_csv_path, pnp_script_path):
          continue
        rel_p = full_p.relative_to(sharepoint_dir).as_posix()
        if single_user_only and "/" in rel_p:
          rel_p = rel_p.split("/", 1)[1]
        write_utf8_zip_bytes(sp_zf, rel_p, full_p.read_bytes())

  # Summarize per-user stats for the Batch Table
  user_summary_map = {}
  for nb in notebooks:
    u_key = nb["owner_email"]
    if u_key not in user_summary_map:
      user_summary_map[u_key] = {
          "gws_email": u_key,
          "m365_upn": nb["m365_upn"],
          "ge_target_email": nb["ge_target_email"],
          "sharepoint_site_url": nb["sharepoint_site_url"],
          "notebook_count": 0,
          "source_count": 0,
          "url_count": 0,
          "notebooks": [],
      }
    user_summary_map[u_key]["notebook_count"] += 1
    user_summary_map[u_key]["source_count"] += len(nb["sources"])
    user_summary_map[u_key]["url_count"] += sum(
        1 for s in nb["sources"] if s.get("source_url")
    )
    user_summary_map[u_key]["notebooks"].append(nb["title"])

  manifest = {
      "user_count": len(discovered_users) or 1,
      "users": sorted(list(discovered_users)),
      "user_summaries": list(user_summary_map.values()),
      "notebook_count": len(notebooks),
      "total_sources": sum(len(n["sources"]) for n in notebooks),
      "total_urls_recovered": sum(
          1 for n in notebooks for s in n["sources"] if s.get("source_url")
      ),
      "total_notes": sum(len(n["notes"]) for n in notebooks),
      "total_artifacts": sum(len(n.get("artifacts", [])) for n in notebooks),
      "sharepoint_zip_path": str(sp_zip_path),
      "spmt_csv_path": str(spmt_csv_path),
      "notebooks": notebooks,
  }
  manifest_path = output_dir / "migration_manifest.json"
  manifest_path.write_text(
      json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
  )
  return manifest


def parse_takeout_zip(
    zip_paths: Path | list[Path],
    output_dir: Path,
    default_owner_email: str = "admin@andresvilla.altostrat.com",
    user_mapping: dict[str, dict] | None = None,
) -> dict:
  """Extracts one or multiple GWS Takeout ZIP files into `extracted_takeout` and builds the SharePoint & GE manifest."""
  extract_dir = output_dir / "extracted_takeout"
  if extract_dir.exists():
    shutil.rmtree(extract_dir)
  extract_dir.mkdir(parents=True, exist_ok=True)

  if isinstance(zip_paths, Path):
    zip_paths = [zip_paths]

  def _extract_single_zip(z_file: Path, dest_root: Path):
    with zipfile.ZipFile(z_file, "r") as zf:
      for info in zf.infolist():
        fname = info.filename
        if not (info.flag_bits & 0x800):
          try:
            fname = fname.encode("cp437").decode("utf-8")
          except Exception:
            pass
        target_path = dest_root / fname
        if info.is_dir() or fname.endswith("/"):
          target_path.mkdir(parents=True, exist_ok=True)
        else:
          target_path.parent.mkdir(parents=True, exist_ok=True)
          with zf.open(info, "r") as src, open(target_path, "wb") as dst:
            shutil.copyfileobj(src, dst)
          if target_path.suffix.lower() == ".zip":
            sub_user = infer_user_email_from_path(
                target_path.stem, default_owner_email
            )
            sub_dest = dest_root / sub_user
            sub_dest.mkdir(parents=True, exist_ok=True)
            try:
              _extract_single_zip(target_path, sub_dest)
              target_path.unlink()
            except Exception:
              pass

  for zp in zip_paths:
    inferred_owner = infer_user_email_from_path(zp.name, default_owner_email)
    dest_sub = (
        extract_dir / inferred_owner if len(zip_paths) > 1 else extract_dir
    )
    dest_sub.mkdir(parents=True, exist_ok=True)
    _extract_single_zip(zp, dest_sub)

  return build_manifest_from_extracted_dir(
      extract_dir=extract_dir,
      output_dir=output_dir,
      default_owner_email=default_owner_email,
      user_mapping=user_mapping,
  )


def sync_gcs_bucket_and_extract(
    gcs_uri: str, output_dir: Path, user_mapping: dict[str, dict] | None = None
) -> dict:
  """Syncs a GWS Customer Takeout GCS bucket (`gs://bucket/prefix/`) and builds the batch manifest."""
  extract_dir = output_dir / "extracted_takeout"
  if extract_dir.exists():
    shutil.rmtree(extract_dir)
  extract_dir.mkdir(parents=True, exist_ok=True)

  proc = subprocess.run(
      ["gcloud", "storage", "cp", "-r", gcs_uri.strip(), str(extract_dir)],
      capture_output=True,
      text=True,
      timeout=300,
  )
  if proc.returncode != 0:
    return {
        "error": (
            f"gcloud storage cp failed (exit {proc.returncode}):"
            f" {proc.stderr[:400]}"
        )
    }

  # Unpack any .zip files downloaded from GCS
  for r, d, files in os.walk(extract_dir):
    for f in files:
      if f.lower().endswith(".zip"):
        zp = Path(r) / f
        user_email = infer_user_email_from_path(str(zp), "user@company.com")
        dest = extract_dir / user_email
        dest.mkdir(parents=True, exist_ok=True)
        try:
          with zipfile.ZipFile(zp, "r") as zf:
            zf.extractall(dest)
          zp.unlink()
        except Exception:
          pass

  return build_manifest_from_extracted_dir(
      extract_dir=extract_dir,
      output_dir=output_dir,
      user_mapping=user_mapping,
  )


def request_with_backoff(
    method: str, url: str, max_retries: int = 4, **kwargs
) -> requests.Response:
  """Executes an HTTP request against discoveryengine.googleapis.com with exponential backoff on 429/5xx."""
  delay = 1.5
  last_resp = None
  for attempt in range(max_retries):
    try:
      resp = requests.request(method, url, **kwargs)
      last_resp = resp
      if resp.status_code not in (429, 500, 502, 503, 504):
        return resp
    except Exception:
      pass
    time.sleep(delay)
    delay *= 2.0
  return last_resp


def get_gcloud_token() -> str:
  try:
    return subprocess.check_output(
        ["gcloud", "auth", "print-access-token"],
        stderr=subprocess.DEVNULL,
        text=True,
        timeout=5,
    ).strip()
  except Exception:
    return ""


def load_checkpoint() -> dict:
  cp_file = WORKSPACE_DIR / "batch_checkpoint.json"
  if cp_file.exists():
    try:
      return json.loads(cp_file.read_text(encoding="utf-8"))
    except Exception:
      return {}
  return {}


def clear_checkpoint() -> None:
  cp_file = WORKSPACE_DIR / "batch_checkpoint.json"
  with CHECKPOINT_LOCK:
    if cp_file.exists():
      cp_file.unlink()


def save_checkpoint_entry(key: str, entry: dict) -> None:
  cp_file = WORKSPACE_DIR / "batch_checkpoint.json"
  with CHECKPOINT_LOCK:
    data = load_checkpoint()
    data[key] = entry
    cp_file.write_text(
        json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def _provision_single_notebook_worker(
    nb: dict,
    endpoint_prefix: str,
    base_parent: str,
    headers: dict,
    project_number: str,
    token: str,
    override_target_email: str,
    use_native_urls: bool,
    resume_checkpoint: bool,
) -> dict:
  """Worker function executed inside ThreadPoolExecutor to migrate one notebook with checkpointing."""
  owner_email = nb.get("owner_email", "default_user")
  nb_title = nb["title"]
  checkpoint_key = f"{owner_email}::{nb_title}"

  if resume_checkpoint:
    cp = load_checkpoint()
    if checkpoint_key in cp and cp[checkpoint_key].get("status") == "completed":
      cached = dict(cp[checkpoint_key])
      cached["resumed_from_checkpoint"] = True
      return cached

  target_email = (
      override_target_email.strip()
      or nb.get("ge_target_email", "").strip()
      or owner_email
  )
  nb_log = {
      "owner_email": owner_email,
      "ge_target_email": target_email,
      "title": nb_title,
      "notebook_id": None,
      "notebook_name": None,
      "sources_uploaded": [],
      "sources_failed": [],
      "notes_created": [],
      "shared_with": None,
      "status": "pending",
  }

  # 1. Create Notebook (?serviceAccountUser=true enables Service Account, Human Admin, and Business User tokens simultaneously)
  create_url = f"{endpoint_prefix}/v1alpha/{base_parent}/notebooks?serviceAccountUser=true"
  resp = request_with_backoff(
      "POST", create_url, headers=headers, json={"title": nb_title}, timeout=25
  )
  if resp is None or resp.status_code not in (200, 201):
    nb_log["status"] = "failed"
    nb_log["error"] = (
        f"CreateNotebook HTTP {resp.status_code if resp else 'ERR'}:"
        f" {resp.text[:300] if resp else 'No response'}"
    )
    return nb_log

  nb_data = resp.json()
  nb_name = nb_data.get("name", "")
  nb_id = (
      nb_name.split("/")[-1] if "/" in nb_name else nb_data.get("notebookId", "")
  )
  nb_log["notebook_id"] = nb_id
  nb_log["notebook_name"] = nb_name

  # 2. Ingest Sources
  batch_url = f"{endpoint_prefix}/v1alpha/{base_parent}/notebooks/{nb_id}/sources:batchCreate"
  upload_url = f"{endpoint_prefix}/upload/v1alpha/{base_parent}/notebooks/{nb_id}/sources:uploadFile"

  for src in nb.get("sources", []):
    src_title = src["title"]
    src_url = src.get("source_url", "")
    yt_url = src.get("youtube_url", "")
    src_type = src.get("source_type", "")
    ingested_via_url = False

    if use_native_urls and (
        yt_url or (src_type == "SOURCE_CONTENT_TYPE_URL" and src_url)
    ):
      user_content = (
          {"videoContent": {"youtubeUrl": yt_url}}
          if yt_url
          else {"webContent": {"url": src_url, "sourceName": src_title}}
      )
      b_resp = request_with_backoff(
          "POST",
          batch_url,
          headers=headers,
          json={"userContents": [user_content]},
          timeout=30,
      )
      if b_resp is not None and b_resp.status_code in (200, 201):
        nb_log["sources_uploaded"].append({
            "title": src_title,
            "method": (
                "batchCreate (youtubeContent)"
                if yt_url
                else "batchCreate (webContent)"
            ),
            "url": yt_url or src_url,
        })
        ingested_via_url = True

    if not ingested_via_url:
      docx_p = Path(src["docx_path"])
      if docx_p.exists():
        file_name_header = sanitize_filename(src_title) + ".docx"
        up_headers = {
            "Authorization": f"Bearer {token}",
            "X-Goog-User-Project": str(project_number),
            "X-Goog-Upload-File-Name": file_name_header,
            "X-Goog-Upload-Protocol": "raw",
            "Content-Type": (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        }
        up_resp = request_with_backoff(
            "POST",
            upload_url,
            headers=up_headers,
            data=docx_p.read_bytes(),
            timeout=60,
        )
        if up_resp is not None and up_resp.status_code in (200, 201):
          nb_log["sources_uploaded"].append({
              "title": src_title,
              "method": "uploadFile (.docx)",
              "url": src_url or None,
          })
        else:
          nb_log["sources_failed"].append({
              "title": src_title,
              "error": (
                  f"HTTP {up_resp.status_code if up_resp else 'ERR'}:"
                  f" {up_resp.text[:200] if up_resp else ''}"
              ),
          })

  # 3. Create Notes
  notes_url = f"{endpoint_prefix}/v1alpha/{base_parent}/notebooks/{nb_id}/notes"
  for note in nb.get("notes", []):
    n_resp = request_with_backoff(
        "POST",
        notes_url,
        headers=headers,
        json={"title": note["title"], "content": note.get("content_text", "")},
        timeout=20,
    )
    if n_resp is not None and n_resp.status_code in (200, 201):
      nb_log["notes_created"].append(note["title"])

  # 4. Share Notebook with Target End-User (`notebooks:share`)
  if target_email and "@" in target_email:
    share_url = f"{endpoint_prefix}/v1alpha/{base_parent}/notebooks:share"
    s_resp = request_with_backoff(
        "POST",
        share_url,
        headers=headers,
        json={
            "name": nb_name,
            "accountAndRoles": [{
                "email": target_email,
                "role": "PROJECT_ROLE_WRITER",
            }],
            "notifyViaEmail": True,
        },
        timeout=20,
    )
    if s_resp is not None and s_resp.status_code in (200, 201):
      nb_log["shared_with"] = target_email

  nb_log["status"] = "completed"
  save_checkpoint_entry(checkpoint_key, nb_log)
  return nb_log


def migrate_to_gemini_enterprise(
    manifest: dict,
    project_number: str,
    location: str = "global",
    access_token: str = "",
    target_owner_email: str = "",
    selected_indices: list[int] | None = None,
    use_native_urls: bool = True,
    max_workers: int = 4,
    resume_checkpoint: bool = True,
) -> dict:
  """Fold 3: Uses a bounded ThreadPoolExecutor + exponential backoff + idempotent

  checkpoint ledger (`batch_checkpoint.json`) to bulk-provision notebooks across
  1 to 100+ users.
  """
  token = access_token.strip() or get_gcloud_token()
  if not token:
    return {
        "status": "error",
        "error": (
            "No OAuthBearer access token provided and `gcloud auth"
            " print-access-token` returned empty."
        ),
    }

  endpoint_prefix = (
      "https://discoveryengine.googleapis.com"
      if location == "global"
      else f"https://{location}-discoveryengine.googleapis.com"
  )
  base_parent = f"projects/{project_number}/locations/{location}"
  headers = {
      "Authorization": f"Bearer {token}",
      "Content-Type": "application/json",
      "X-Goog-User-Project": str(project_number),
  }

  try:
    requests.post(
        f"{endpoint_prefix}/v1alpha/{base_parent}/accounts/me:getOrCreate",
        headers=headers,
        json={},
        timeout=15,
    )
  except Exception:
    pass

  notebooks = manifest.get("notebooks", [])
  tasks_to_run = [
      nb
      for idx, nb in enumerate(notebooks)
      if (selected_indices is None or idx in selected_indices)
  ]

  results = []
  workers = max(1, min(int(max_workers or 4), 12))
  with ThreadPoolExecutor(max_workers=workers) as pool:
    futures = [
        pool.submit(
            _provision_single_notebook_worker,
            nb,
            endpoint_prefix,
            base_parent,
            headers,
            project_number,
            token,
            target_owner_email,
            use_native_urls,
            resume_checkpoint,
        )
        for nb in tasks_to_run
    ]
    for fut in as_completed(futures):
      results.append(fut.result())

  completed_count = sum(1 for r in results if r.get("status") == "completed")
  failed_count = len(results) - completed_count
  return {
      "status": "success" if failed_count == 0 else "partial_success",
      "workers_used": workers,
      "resume_checkpoint_enabled": resume_checkpoint,
      "total_processed": len(results),
      "completed_count": completed_count,
      "failed_count": failed_count,
      "migrated_notebooks": results,
  }


DEFAULT_CSV_TEMPLATE = """gws_email,m365_upn,ge_target_email,sharepoint_site_url
admin@andresvilla.altostrat.com,admin@contoso-m365.com,admin@andresvilla.altostrat.com,https://contoso.sharepoint.com/sites/GeminiNotebooks/Shared Documents/admin@contoso-m365.com
maria.lopez@andresvilla.altostrat.com,maria.lopez@contoso-m365.com,maria.lopez@andresvilla.altostrat.com,https://contoso.sharepoint.com/sites/GeminiNotebooks/Shared Documents/maria.lopez@contoso-m365.com
carlos.ruiz@andresvilla.altostrat.com,carlos.ruiz@contoso-m365.com,carlos.ruiz@andresvilla.altostrat.com,https://contoso.sharepoint.com/sites/GeminiNotebooks/Shared Documents/carlos.ruiz@contoso-m365.com"""


WORKBENCH_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>GWS NotebookLM → SharePoint & Gemini Enterprise Migrator</title>
  <style>
    :root {
      --bg: #f8f9fa;
      --card: #ffffff;
      --primary: #1a73e8;
      --primary-dark: #1557b0;
      --success: #137333;
      --success-bg: #e6f4ea;
      --amber: #b06000;
      --amber-bg: #fef7e0;
      --purple: #681da8;
      --purple-bg: #f3e8fd;
      --danger: #c5221f;
      --danger-bg: #fce8e6;
      --border: #dadce0;
      --text: #202124;
      --muted: #5f6368;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: 'Google Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.5;
    }
    header {
      background: #1a73e8;
      color: #fff;
      padding: 20px 32px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      box-shadow: 0 2px 6px rgba(0,0,0,0.12);
    }
    header h1 { margin: 0; font-size: 20px; font-weight: 600; }
    header p { margin: 4px 0 0; font-size: 13px; opacity: 0.9; }
    .container {
      max-width: 1280px;
      margin: 24px auto;
      padding: 0 24px 64px;
    }
    .tabs {
      display: flex;
      gap: 10px;
      margin-bottom: 20px;
      border-bottom: 2px solid var(--border);
      padding-bottom: 10px;
    }
    .tab-btn {
      background: #fff;
      color: var(--muted);
      border: 1px solid var(--border);
      padding: 10px 20px;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
    }
    .tab-btn.active {
      background: var(--primary);
      color: #fff;
      border-color: var(--primary);
    }
    .grid-folds {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 16px;
      margin-bottom: 24px;
    }
    .fold-pill {
      background: var(--card);
      border: 1px solid var(--border);
      border-top: 4px solid var(--primary);
      border-radius: 8px;
      padding: 14px 16px;
    }
    .fold-pill h3 { margin: 0 0 6px; font-size: 14px; color: var(--primary); text-transform: uppercase; letter-spacing: 0.5px; }
    .fold-pill p { margin: 0; font-size: 13px; color: var(--muted); }
    .card {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 24px;
      margin-bottom: 24px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .card h2 { margin-top: 0; font-size: 18px; display: flex; align-items: center; gap: 10px; }
    .badge {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 3px 9px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 600;
      white-space: nowrap;
    }
    .badge-blue { background: #e8f0fe; color: var(--primary); }
    .badge-green { background: var(--success-bg); color: var(--success); }
    .badge-amber { background: var(--amber-bg); color: var(--amber); }
    .badge-purple { background: var(--purple-bg); color: var(--purple); }
    .badge-red { background: var(--danger-bg); color: var(--danger); }
    .btn {
      background: var(--primary);
      color: #fff;
      border: none;
      padding: 10px 18px;
      border-radius: 6px;
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      text-decoration: none;
    }
    .btn:hover { background: var(--primary-dark); }
    .btn-success { background: var(--success); }
    .btn-success:hover { background: #0d5324; }
    .btn-purple { background: var(--purple); }
    .btn-outline {
      background: #fff;
      color: var(--text);
      border: 1px solid var(--border);
    }
    input[type="text"], input[type="password"], input[type="file"], input[type="number"], select, textarea {
      width: 100%;
      padding: 9px 12px;
      border: 1px solid var(--border);
      border-radius: 6px;
      font-size: 13px;
      margin-top: 4px;
      font-family: inherit;
    }
    textarea { font-family: monospace; font-size: 12px; }
    .form-row {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 16px;
      margin-bottom: 16px;
    }
    label { font-size: 13px; font-weight: 600; color: var(--text); display: block; }
    .nb-card {
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 16px;
      margin-bottom: 16px;
      background: #fafbfc;
    }
    .nb-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }
    .src-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 6px;
      overflow: hidden;
    }
    .src-table th {
      background: #f1f3f4;
      text-align: left;
      padding: 8px 12px;
      font-size: 12px;
      color: var(--muted);
      border-bottom: 1px solid var(--border);
    }
    .src-table td {
      padding: 8px 12px;
      border-bottom: 1px solid #f1f3f4;
      vertical-align: middle;
    }
    .src-table tr:last-child td { border-bottom: none; }
    .url-link {
      color: var(--primary);
      text-decoration: none;
      font-family: monospace;
      font-size: 12px;
      word-break: break-all;
    }
    .url-link:hover { text-decoration: underline; }
    pre {
      background: #202124;
      color: #e8eaed;
      padding: 14px;
      border-radius: 8px;
      font-size: 12px;
      overflow-x: auto;
    }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>🚀 GWS NotebookLM → SharePoint & Gemini Enterprise Migrator</h1>
      <p>Functional 3-Fold Pipeline • Single-User & Multi-User Batch Wave Engine (Multi-ZIP • GCS Bucket Sync • SharePoint SPMT/PnP • Parallel GE Provisioning)</p>
    </div>
    <div style="display:flex; gap:8px;">
      <span class="badge badge-green">UTF-8 ZIP Fix Active (0x800)</span>
      <span class="badge badge-purple">Multi-User Batch Engine Active</span>
    </div>
  </header>

  <div class="container">
    <div class="tabs">
      <button id="tabSingleBtn" class="tab-btn active" onclick="switchTab('single')">👤 Fold 1–3: Notebook & Source Inspector + GE Provisioner</button>
      <button id="tabBatchBtn" class="tab-btn" onclick="switchTab('batch')">🏢 Multi-User / 100-User Batch Wave Ingestion & Mapping</button>
    </div>

    <!-- TAB 2: FUNCTIONAL MULTI-USER / 100-USER BATCH ENGINE -->
    <div id="tabBatchContent" style="display:none;">
      <div class="card">
        <h2>🏢 Step 1: Configure Multi-User Batch Identity Mapping (`user_mapping.csv`)</h2>
        <p style="color:var(--muted); font-size:13px; margin-top:0;">
          Paste or edit your batch user mapping CSV (`gws_email,m365_upn,ge_target_email,sharepoint_site_url`) or upload a `.csv` file. You can also auto-generate a N-user CSV wave from your domain template and immediately build the physical `.docx` SharePoint folders and SPMT manifest on disk.
        </p>
        <div style="display:flex; gap:10px; margin-bottom:10px; flex-wrap:wrap; align-items:center;">
          <input type="file" id="csvFileInput" accept=".csv" style="max-width:260px; margin:0;" onchange="loadCsvFile(this)" />
          <button class="btn btn-outline" onclick="populateNUsersCsv(10)">➕ Populate 10-User Wave CSV</button>
          <button class="btn btn-outline" onclick="populateNUsersCsv(100)">➕ Populate 100-User Wave CSV</button>
          <button class="btn btn-purple" onclick="applyBatchMappingToExtracted()">⚡ Apply CSV Mapping & Build Physical SharePoint Package + SPMT Manifest</button>
        </div>
        <textarea id="batchCsvText" rows="7">__DEFAULT_CSV_TEMPLATE__</textarea>
      </div>

      <div class="card">
        <h2>📦 Step 2: Batch Source Ingestion (Upload Multiple User Takeout `.zip` Files OR Sync from GWS Customer Takeout GCS Bucket)</h2>
        <div class="form-row" style="grid-template-columns: 1fr 1fr;">
          <div style="border:1px solid var(--border); padding:16px; border-radius:8px;">
            <label>Option A: Upload Multiple User Takeout `.zip` Archives (`takeout-user1@domain.zip`, ...)</label>
            <input type="file" id="multiZipInput" accept=".zip" multiple style="margin:10px 0;" />
            <button class="btn" onclick="uploadMultiZipBatch()">📂 Extract All Selected User `.zip` Files with CSV Mapping</button>
          </div>
          <div style="border:1px solid var(--border); padding:16px; border-radius:8px;">
            <label>Option B: Pull Directly from GWS Customer Takeout GCS Bucket (`gs://bucket/prefix/`)</label>
            <input type="text" id="gcsBucketInput" placeholder="gs://gws-customer-takeout-bucket/export_wave_01/" style="margin:10px 0;" />
            <button class="btn" onclick="syncGcsBucketBatch()">☁️ Sync GCS Bucket & Build Multi-User Package</button>
          </div>
        </div>
      </div>

      <div class="card">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
          <h2 style="margin:0;">📊 Step 3: Active Batch Wave Users & SharePoint / GE Provisioning Queue</h2>
          <div style="display:flex; gap:10px;">
            <a href="/api/download-sharepoint-zip" class="btn btn-success">📥 Download Multi-User `SharePoint_Ready_Notebooks.zip` (+ SPMT CSV & PnP Script)</a>
            <button class="btn" onclick="switchTab('single'); runGeMigration();">🚀 Execute Parallel GE Provisioning for All Mapped Users</button>
          </div>
        </div>
        <table class="src-table">
          <thead>
            <tr>
              <th>#</th>
              <th>GWS Source Owner</th>
              <th>Target M365 / SharePoint UPN</th>
              <th>Target Gemini Enterprise Email (`notebooks:share`)</th>
              <th>SharePoint Destination URL</th>
              <th>Notebooks</th>
              <th>Sources (`.docx`)</th>
              <th>Recovered URLs</th>
            </tr>
          </thead>
          <tbody id="activeBatchUsersBody">
            <tr><td colspan="8">Loading active user batch manifest...</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- TAB 1: SINGLE / MULTI-USER NOTEBOOK INSPECTOR & GE PROVISIONER -->
    <div id="tabSingleContent">
      <div class="grid-folds">
        <div class="fold-pill">
          <h3>Fold 1 • GWS Takeout & URL Extraction</h3>
          <p>Unpacks <code>Takeout/NotebookLM/*.zip</code> (single or multi-user), deduplicates notebooks, and recovers original <strong>Website URLs</strong> & <strong>YouTube URLs</strong>.</p>
        </div>
        <div class="fold-pill">
          <h3>Fold 2 • SharePoint / OneDrive Staging</h3>
          <p>Converts sources into clean <strong>Microsoft Word (.docx)</strong> files + <code>sharepoint_spmt_manifest.csv</code> & <code>upload_to_sharepoint_pnp.ps1</code>.</p>
        </div>
        <div class="fold-pill">
          <h3>Fold 3 • Gemini Enterprise Bulk Creation</h3>
          <p>Parallel worker pool + <code>batch_checkpoint.json</code> idempotency ledger calling <code>discoveryengine.googleapis.com/v1alpha</code>.</p>
        </div>
      </div>

      <div class="card">
        <h2>📂 Fold 1 & 2: Upload GWS Takeout ZIP(s) → Extract Original URLs & SharePoint (.docx) Package</h2>
        <div style="display:flex; gap:12px; align-items:flex-end; flex-wrap:wrap;">
          <div style="flex:1; min-width:280px;">
            <label>Select 1 or Multiple GWS Takeout ZIP files (<code>takeout-*.zip</code>)</label>
            <input type="file" id="zipFile" accept=".zip" multiple />
          </div>
          <button class="btn" onclick="uploadTakeout()">⚡ Extract & Convert to SharePoint (.docx)</button>
          <a id="downloadSpBtn" href="/api/download-sharepoint-zip" class="btn btn-success" style="display:none;">
            📥 Download Clean SharePoint_Ready_Notebooks.zip (UTF-8 + SPMT CSV)
          </a>
        </div>

        <div id="manifestContainer" style="margin-top:20px; display:none;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
            <h3 style="margin:0; font-size:16px;">Extracted Notebooks, Original URLs & SharePoint Staging</h3>
            <div style="display:flex; gap:8px;">
              <span id="manifestStats" class="badge badge-blue"></span>
              <span id="urlStats" class="badge badge-green"></span>
            </div>
          </div>
          <div id="notebookList"></div>
        </div>
      </div>

      <div class="card">
        <h2>☁️ Fold 3: Bulk-Create Notebooks on Gemini Enterprise (Parallel Worker Pool + Checkpoint Ledger)</h2>
        <div class="form-row">
          <div>
            <label>GCP Project Number (Required)</label>
            <input type="text" id="projectNumber" placeholder="e.g. 123456789012" />
          </div>
          <div>
            <label>NotebookLM Enterprise Location</label>
            <select id="location">
              <option value="global">global (https://discoveryengine.googleapis.com)</option>
              <option value="us">us (https://us-discoveryengine.googleapis.com)</option>
              <option value="eu">eu (https://eu-discoveryengine.googleapis.com)</option>
            </select>
          </div>
          <div>
            <label>Parallel Worker Threads (1–12)</label>
            <input type="number" id="maxWorkers" value="4" min="1" max="12" />
          </div>
        </div>
        <div class="form-row">
          <div>
            <label>Target End-User Email Override (Leave blank to use each user's mapped `ge_target_email`)</label>
            <input type="text" id="targetEmail" placeholder="Optional override (otherwise uses per-notebook ge_target_email)" />
          </div>
          <div>
            <label>Source Ingestion Mode for Website & YouTube URLs</label>
            <select id="useNativeUrls">
              <option value="true">Hybrid (Live URLs via batchCreate + .docx via uploadFile)</option>
              <option value="false">All SharePoint .docx Snapshots (uploadFile)</option>
            </select>
          </div>
          <div>
            <label>OAuth 2.0 Bearer Token (Optional if <code>gcloud auth</code> active)</label>
            <input type="password" id="accessToken" placeholder="Paste OAuth2 Bearer Token" />
          </div>
        </div>
        <div style="display:flex; gap:10px; align-items:center;">
          <button class="btn" onclick="runGeMigration()">🚀 Bulk-Create Selected Notebooks on Gemini Enterprise</button>
          <button class="btn btn-outline" onclick="clearCheckpointState()">🗑️ Reset Checkpoint Ledger (`batch_checkpoint.json`)</button>
        </div>

        <div id="migrationResultBox" style="margin-top:20px; display:none;">
          <h3 style="margin:0 0 8px; font-size:15px;">Gemini Enterprise Provisioning Log</h3>
          <pre id="migrationOutput"></pre>
        </div>
      </div>
    </div>
  </div>

  <script>
    let currentManifest = null;

    function switchTab(mode) {
      document.getElementById('tabSingleContent').style.display = (mode === 'single') ? 'block' : 'none';
      document.getElementById('tabBatchContent').style.display = (mode === 'batch') ? 'block' : 'none';
      document.getElementById('tabSingleBtn').classList.toggle('active', mode === 'single');
      document.getElementById('tabBatchBtn').classList.toggle('active', mode === 'batch');
    }

    function loadCsvFile(input) {
      if (!input.files.length) return;
      const reader = new FileReader();
      reader.onload = e => {
        document.getElementById('batchCsvText').value = e.target.result;
      };
      reader.readAsText(input.files[0]);
    }

    function populateNUsersCsv(count) {
      const lines = ['gws_email,m365_upn,ge_target_email,sharepoint_site_url'];
      for (let i = 1; i <= count; i++) {
        const idx = String(i).padStart(3, '0');
        const gws = `user${idx}@andresvilla.altostrat.com`;
        const m365 = `user${idx}@contoso-m365.com`;
        const ge = `user${idx}@andresvilla.altostrat.com`;
        const sp = `https://contoso.sharepoint.com/sites/GeminiNotebooks/Shared Documents/${m365}`;
        lines.push(`${gws},${m365},${ge},${sp}`);
      }
      document.getElementById('batchCsvText').value = lines.join('\\n');
    }

    async function applyBatchMappingToExtracted() {
      const csvText = document.getElementById('batchCsvText').value;
      const resp = await fetch('/api/apply-batch-mapping', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({csv_text: csvText})
      });
      const data = await resp.json();
      if (data.error) {
        alert('Error: ' + data.error);
        return;
      }
      currentManifest = data;
      renderManifest(data);
      alert(`Successfully built physical SharePoint .docx folders + SPMT CSV for ${data.user_count} user(s) (${data.notebook_count} notebooks, ${data.total_sources} sources)!`);
    }

    async function uploadMultiZipBatch() {
      const fileInput = document.getElementById('multiZipInput');
      if (!fileInput.files.length) {
        alert('Please select one or more user Takeout .zip files.');
        return;
      }
      const formData = new FormData();
      formData.append('csv_text', document.getElementById('batchCsvText').value);
      for (let i = 0; i < fileInput.files.length; i++) {
        formData.append('files', fileInput.files[i]);
      }
      const resp = await fetch('/api/extract-takeout', {
        method: 'POST',
        body: formData
      });
      const data = await resp.json();
      if (data.error) {
        alert('Error: ' + data.error);
        return;
      }
      currentManifest = data;
      renderManifest(data);
    }

    async function syncGcsBucketBatch() {
      const gcsUri = document.getElementById('gcsBucketInput').value.trim();
      if (!gcsUri.startsWith('gs://')) {
        alert('Please enter a valid GCS URI starting with gs://');
        return;
      }
      const csvText = document.getElementById('batchCsvText').value;
      const resp = await fetch('/api/sync-gcs-batch', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({gcs_uri: gcsUri, csv_text: csvText})
      });
      const data = await resp.json();
      if (data.error) {
        alert('GCS Sync Error: ' + data.error);
        return;
      }
      currentManifest = data;
      renderManifest(data);
    }

    function formatSourceBadge(stype) {
      if (stype === 'SOURCE_CONTENT_TYPE_URL') {
        return '<span class="badge badge-blue">🌐 Website URL</span>';
      } else if (stype === 'SOURCE_CONTENT_TYPE_YOUTUBE_VIDEO') {
        return '<span class="badge badge-amber">▶️ YouTube Video</span>';
      } else if (stype === 'SOURCE_CONTENT_TYPE_PDF') {
        return '<span class="badge badge-purple">📕 PDF Document</span>';
      } else if (stype === 'SOURCE_CONTENT_TYPE_POWERPOINT') {
        return '<span class="badge badge-amber">📊 PowerPoint</span>';
      }
      return `<span class="badge badge-blue">📄 ${stype.replace('SOURCE_CONTENT_TYPE_', '')}</span>`;
    }

    async function uploadTakeout() {
      const fileInput = document.getElementById('zipFile');
      if (!fileInput.files.length) {
        alert('Please select at least one GWS Takeout .zip file.');
        return;
      }
      const formData = new FormData();
      for (let i = 0; i < fileInput.files.length; i++) {
        formData.append('files', fileInput.files[i]);
      }
      const resp = await fetch('/api/extract-takeout', {
        method: 'POST',
        body: formData
      });
      const data = await resp.json();
      if (data.error) {
        alert('Error: ' + data.error);
        return;
      }
      currentManifest = data;
      renderManifest(data);
    }

    function renderManifest(manifest) {
      document.getElementById('manifestContainer').style.display = 'block';
      document.getElementById('downloadSpBtn').style.display = 'inline-flex';
      document.getElementById('manifestStats').textContent =
        `${manifest.user_count || 1} User(s) • ${manifest.notebook_count} Notebooks • ${manifest.total_sources} Sources • ${manifest.total_artifacts || 0} Studio Artifacts`;
      document.getElementById('urlStats').textContent =
        `🔗 ${manifest.total_urls_recovered || 0} Original Web/YouTube URLs Recovered`;

      // Render Active Batch Users Table in Tab 2
      const batchTbody = document.getElementById('activeBatchUsersBody');
      const summaries = manifest.user_summaries || [];
      batchTbody.innerHTML = summaries.map((u, idx) => `
        <tr>
          <td><span class="badge badge-purple">${idx + 1}</span></td>
          <td style="font-family:monospace; font-size:12px;">${u.gws_email}</td>
          <td style="font-family:monospace; font-size:12px;">${u.m365_upn}</td>
          <td style="font-family:monospace; font-size:12px; color:var(--primary); font-weight:600;">${u.ge_target_email}</td>
          <td style="font-family:monospace; font-size:11px; color:var(--muted);">${u.sharepoint_site_url}</td>
          <td><span class="badge badge-blue">${u.notebook_count} Notebooks</span></td>
          <td><span class="badge badge-green">${u.source_count} .docx</span></td>
          <td><span class="badge badge-amber">${u.url_count} Live URLs</span></td>
        </tr>
      `).join('') || '<tr><td colspan="8">No users extracted yet.</td></tr>';

      // Render Notebook Cards in Tab 1 (cap detailed DOM cards at 60 to keep browser snappy on 100-user batches)
      const listEl = document.getElementById('notebookList');
      listEl.innerHTML = '';
      const displayNbs = manifest.notebooks.slice(0, 60);
      displayNbs.forEach((nb, idx) => {
        const rowsHtml = nb.sources.map((s) => {
          let originCell = '<span style="color:var(--muted); font-size:12px;">📁 Uploaded File Snapshot (Converted to .docx)</span>';
          if (s.source_url) {
            const extra = s.youtube_channel ? ` <span class="badge badge-amber" style="margin-left:6px;">📺 ${s.youtube_channel}</span>` : '';
            originCell = `<a class="url-link" href="${s.source_url}" target="_blank" rel="noopener">🔗 ${s.source_url}</a>${extra}`;
          }
          return `
            <tr>
              <td style="font-weight:500;">${s.title}</td>
              <td>${formatSourceBadge(s.source_type)}</td>
              <td>${originCell}</td>
              <td style="font-family:monospace; font-size:12px; color:var(--muted);">${s.relative_sharepoint_path}</td>
            </tr>
          `;
        }).join('');

        const artifactsHtml = (nb.artifacts && nb.artifacts.length)
          ? `<div style="margin-top:10px; font-size:12px; color:var(--muted);">✨ <strong>Studio Artifacts (${nb.artifacts.length}):</strong> ` +
            nb.artifacts.map(a => `<span class="badge badge-purple" style="margin-right:6px;">${a.title} (${a.artifact_type.replace('ARTIFACT_TYPE_', '').replace('APP_TYPE_', '')})</span>`).join('') +
            `</div>`
          : '';

        const div = document.createElement('div');
        div.className = 'nb-card';
        div.innerHTML = `
          <div class="nb-header">
            <label style="font-size:16px; display:flex; align-items:center; gap:8px; cursor:pointer; margin:0;">
              <input type="checkbox" class="nb-check" value="${idx}" checked />
              <span>${nb.emoji || '📓'} <strong>${nb.title}</strong></span>
            </label>
            <div style="display:flex; gap:8px;">
              <span class="badge badge-purple">👤 GWS: ${nb.owner_email} → GE: ${nb.ge_target_email}</span>
              <span class="badge badge-blue">SharePoint: ${nb.sharepoint_folder}/</span>
              <span class="badge badge-green">${nb.sources.length} Sources</span>
            </div>
          </div>
          <table class="src-table">
            <thead>
              <tr>
                <th style="width:28%;">Source Title</th>
                <th style="width:14%;">Source Type</th>
                <th style="width:38%;">Original Source URL / Origin</th>
                <th style="width:20%;">SharePoint (.docx) Path</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml || '<tr><td colspan="4">No sources found</td></tr>'}
            </tbody>
          </table>
          ${artifactsHtml}
        `;
        listEl.appendChild(div);
      });

      if (manifest.notebooks.length > 60) {
        const moreNote = document.createElement('div');
        moreNote.style.padding = '12px';
        moreNote.style.textAlign = 'center';
        moreNote.style.color = 'var(--muted)';
        moreNote.innerHTML = `Showing first 60 of <strong>${manifest.notebooks.length}</strong> notebooks in detail view. All <strong>${manifest.notebooks.length}</strong> notebooks are staged on disk and included in <code>SharePoint_Ready_Notebooks.zip</code> and GE Batch Provisioning.`;
        listEl.appendChild(moreNote);
      }
    }

    async function clearCheckpointState() {
      await fetch('/api/clear-checkpoint', {method: 'POST'});
      alert('Checkpoint ledger (batch_checkpoint.json) cleared.');
    }

    async function runGeMigration() {
      if (!currentManifest) {
        alert('Please extract a Takeout ZIP first.');
        return;
      }
      const projectNumber = document.getElementById('projectNumber').value.trim();
      if (!projectNumber) {
        alert('Please enter your numeric GCP Project Number.');
        return;
      }
      const location = document.getElementById('location').value;
      const targetEmail = document.getElementById('targetEmail').value.trim();
      const accessToken = document.getElementById('accessToken').value.trim();
      const useNativeUrls = document.getElementById('useNativeUrls').value === 'true';
      const maxWorkers = parseInt(document.getElementById('maxWorkers').value || '4');

      const outBox = document.getElementById('migrationResultBox');
      const outPre = document.getElementById('migrationOutput');
      outBox.style.display = 'block';
      outPre.textContent = `Running Discovery Engine v1alpha parallel migration (${maxWorkers} workers across ${currentManifest.notebook_count} notebooks)...`;

      const resp = await fetch('/api/migrate-ge', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
          project_number: projectNumber,
          location: location,
          target_owner_email: targetEmail,
          access_token: accessToken,
          selected_indices: null,
          use_native_urls: useNativeUrls,
          max_workers: maxWorkers
        })
      });
      const resData = await resp.json();
      outPre.textContent = JSON.stringify(resData, null, 2);
    }

    fetch('/api/current-manifest').then(r => r.json()).then(d => {
      if (d && d.notebook_count !== undefined) {
        currentManifest = d;
        renderManifest(d);
      }
    }).catch(() => {});
  </script>
</body>
</html>
""".replace("__DEFAULT_CSV_TEMPLATE__", DEFAULT_CSV_TEMPLATE)


def parse_multipart_form_data(content_type: str, body: bytes) -> tuple[dict[str, str], list[tuple[str, bytes]]]:
  """Parses multipart/form-data returning (fields_dict, list_of_(filename, file_bytes))."""
  fields = {}
  files = []
  if "boundary=" not in content_type:
    return fields, [("takeout_uploaded.zip", body)]

  msg_bytes = (
      f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode(
          "utf-8"
      )
      + body
  )
  msg = email.parser.BytesParser().parsebytes(msg_bytes)
  for part in msg.walk():
    if part.is_multipart():
      continue
    disp = part.get("Content-Disposition", "")
    fname = part.get_filename()
    payload = part.get_payload(decode=True) or b""
    name_match = re.search(r'name="([^"]+)"', disp)
    field_name = name_match.group(1) if name_match else ""
    if fname:
      files.append((fname, payload))
    elif field_name:
      fields[field_name] = payload.decode("utf-8", errors="replace")
  return fields, files


class WorkbenchHandler(BaseHTTPRequestHandler):

  def log_message(self, format, *args):
    pass

  def _send_json(self, data: dict, status: int = 200):
    raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
    self.send_response(status)
    self.send_header("Content-Type", "application/json; charset=utf-8")
    self.send_header("Content-Length", str(len(raw)))
    self.end_headers()
    self.wfile.write(raw)

  def do_GET(self):
    if self.path == "/" or self.path.startswith("/?"):
      raw = WORKBENCH_HTML.encode("utf-8")
      self.send_response(200)
      self.send_header("Content-Type", "text/html; charset=utf-8")
      self.send_header("Content-Length", str(len(raw)))
      self.end_headers()
      self.wfile.write(raw)
    elif self.path == "/api/current-manifest":
      m_path = WORKSPACE_DIR / "migration_manifest.json"
      if m_path.exists():
        self._send_json(json.loads(m_path.read_text(encoding="utf-8")))
      else:
        self._send_json({})
    elif self.path == "/api/download-sharepoint-zip":
      sp_zip = WORKSPACE_DIR / "SharePoint_Ready_Notebooks.zip"
      if not sp_zip.exists():
        self._send_json({"error": "No SharePoint ZIP generated yet"}, 404)
        return
      raw = sp_zip.read_bytes()
      self.send_response(200)
      self.send_header("Content-Type", "application/zip")
      self.send_header(
          "Content-Disposition",
          'attachment; filename="SharePoint_Ready_Notebooks.zip"',
      )
      self.send_header("Content-Length", str(len(raw)))
      self.end_headers()
      self.wfile.write(raw)
    else:
      self.send_response(404)
      self.end_headers()

  def do_POST(self):
    content_len = int(self.headers.get("Content-Length", "0"))
    body = self.rfile.read(content_len) if content_len > 0 else b""

    if self.path == "/api/extract-takeout":
      content_type = self.headers.get("Content-Type", "")
      fields, uploaded_files = parse_multipart_form_data(content_type, body)
      user_map = parse_user_mapping_csv_text(fields.get("csv_text", ""))

      uploads_dir = WORKSPACE_DIR / "uploaded_zips"
      if uploads_dir.exists():
        shutil.rmtree(uploads_dir)
      uploads_dir.mkdir(parents=True, exist_ok=True)

      saved_zips = []
      for idx, (fname, fbytes) in enumerate(uploaded_files):
        safe_z = uploads_dir / sanitize_filename(fname or f"takeout_{idx}.zip")
        safe_z.write_bytes(fbytes)
        saved_zips.append(safe_z)
        if idx == 0:
          (WORKSPACE_DIR / "takeout_uploaded.zip").write_bytes(fbytes)

      if not saved_zips:
        self._send_json({"error": "No valid .zip files uploaded."}, 400)
        return

      try:
        manifest = parse_takeout_zip(
            saved_zips, WORKSPACE_DIR, user_mapping=user_map
        )
        self._send_json(manifest)
      except Exception as e:
        self._send_json({"error": str(e)}, 500)

    elif self.path == "/api/apply-batch-mapping":
      req_data = json.loads(body.decode("utf-8") or "{}")
      user_map = parse_user_mapping_csv_text(req_data.get("csv_text", ""))
      extract_dir = WORKSPACE_DIR / "extracted_takeout"
      if not extract_dir.exists():
        self._send_json(
            {"error": "Please upload at least one Takeout ZIP first."}, 400
        )
        return
      try:
        manifest = build_manifest_from_extracted_dir(
            extract_dir=extract_dir,
            output_dir=WORKSPACE_DIR,
            user_mapping=user_map,
        )
        self._send_json(manifest)
      except Exception as e:
        self._send_json({"error": str(e)}, 500)

    elif self.path == "/api/sync-gcs-batch":
      req_data = json.loads(body.decode("utf-8") or "{}")
      gcs_uri = req_data.get("gcs_uri", "")
      user_map = parse_user_mapping_csv_text(req_data.get("csv_text", ""))
      res = sync_gcs_bucket_and_extract(
          gcs_uri, WORKSPACE_DIR, user_mapping=user_map
      )
      self._send_json(res, 400 if "error" in res else 200)

    elif self.path == "/api/clear-checkpoint":
      clear_checkpoint()
      self._send_json({"status": "cleared"})

    elif self.path == "/api/migrate-ge":
      req_data = json.loads(body.decode("utf-8") or "{}")
      m_path = WORKSPACE_DIR / "migration_manifest.json"
      if not m_path.exists():
        self._send_json(
            {"error": "No migration_manifest.json found. Extract ZIP first."},
            400,
        )
        return
      manifest = json.loads(m_path.read_text(encoding="utf-8"))
      res = migrate_to_gemini_enterprise(
          manifest=manifest,
          project_number=req_data.get("project_number", ""),
          location=req_data.get("location", "global"),
          access_token=req_data.get("access_token", ""),
          target_owner_email=req_data.get("target_owner_email", ""),
          selected_indices=req_data.get("selected_indices"),
          use_native_urls=req_data.get("use_native_urls", True),
          max_workers=req_data.get("max_workers", 4),
      )
      self._send_json(res)


def main():
  parser = argparse.ArgumentParser(
      description=(
          "GWS NotebookLM to SharePoint & Gemini Enterprise Migrator (Single &"
          " Multi-User Batch Engine)"
      )
  )
  parser.add_argument(
      "--extract-zip",
      type=str,
      nargs="+",
      help="One or more GWS Takeout ZIP files",
  )
  parser.add_argument(
      "--gcs-bucket",
      type=str,
      help=(
          "GCS URI from GWS Customer Takeout (e.g."
          " gs://my-takeout-bucket/export_prefix/)"
      ),
  )
  parser.add_argument(
      "--user-mapping-csv",
      type=str,
      help=(
          "CSV mapping gws_email,m365_upn,ge_target_email,sharepoint_site_url"
      ),
  )
  parser.add_argument(
      "--migrate-ge-project",
      type=str,
      help="Numeric GCP Project Number to run Fold 3 bulk provisioning",
  )
  parser.add_argument(
      "--location",
      type=str,
      default="global",
      help="GE location (global, us, eu)",
  )
  parser.add_argument(
      "--max-workers",
      type=int,
      default=4,
      help="Concurrent threads for Fold 3 GE provisioning",
  )
  parser.add_argument(
      "--serve", action="store_true", help="Run Web Workbench UI"
  )
  parser.add_argument(
      "--port", type=int, default=8765, help="Port for Web Workbench"
  )
  args = parser.parse_args()

  user_map = {}
  if args.user_mapping_csv and Path(args.user_mapping_csv).exists():
    user_map = parse_user_mapping_csv_text(
        Path(args.user_mapping_csv).read_text(encoding="utf-8")
    )

  if args.gcs_bucket:
    manifest = sync_gcs_bucket_and_extract(
        args.gcs_bucket, WORKSPACE_DIR, user_mapping=user_map
    )
    print(json.dumps({"status": "gcs_extracted", "users": manifest.get("user_count"), "notebooks": manifest.get("notebook_count")}, indent=2))

  if args.extract_zip:
    z_paths = [Path(p) for p in args.extract_zip]
    manifest = parse_takeout_zip(
        z_paths, WORKSPACE_DIR, user_mapping=user_map
    )
    print(
        f"Extracted {manifest['user_count']} user(s),"
        f" {manifest['notebook_count']} notebooks,"
        f" {manifest['total_sources']} sources,"
        f" {manifest['total_urls_recovered']} original URLs recovered."
    )
    print(f"SharePoint ZIP: {manifest['sharepoint_zip_path']}")
    print(f"SPMT Manifest CSV: {manifest['spmt_csv_path']}")

  if args.migrate_ge_project:
    m_path = WORKSPACE_DIR / "migration_manifest.json"
    if m_path.exists():
      manifest = json.loads(m_path.read_text(encoding="utf-8"))
      res = migrate_to_gemini_enterprise(
          manifest=manifest,
          project_number=args.migrate_ge_project,
          location=args.location,
          max_workers=args.max_workers,
      )
      print(json.dumps(res, indent=2))

  if args.serve:
    server = HTTPServer(("0.0.0.0", args.port), WorkbenchHandler)
    print(
        "Workbench running on"
        f" http://andresvrlin.c.googlers.com:{args.port}"
    )
    server.serve_forever()


if __name__ == "__main__":
  main()
