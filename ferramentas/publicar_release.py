"""Build a data-free portable package; optionally publish it to GitHub Releases."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import subprocess
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from atualizacao.atualizador import (  # noqa: E402
    APP_FILES, APP_ID, MANIFEST, UpdateError, read_version, sha256, version_tuple, write_json,
)


def build_package(root: Path, version: str | None = None, output: Path | None = None) -> Path:
    info = read_version(root)
    version = version or info["version"]
    if version_tuple(version) < version_tuple(info["version"]):
        raise UpdateError("A nova versao nao pode ser menor que a atual.")
    info["version"] = version
    payloads = {}
    for name in APP_FILES:
        path = root / name
        if path.is_symlink() or not path.is_file():
            raise UpdateError(f"Arquivo obrigatorio ausente ou invalido: {name}")
        payloads[name] = path.read_bytes()
        if name.endswith(".py"):
            compile(payloads[name], name, "exec")
    payloads["versao.json"] = json.dumps(info, ensure_ascii=True, indent=2).encode("utf-8")
    manifest = {"schema": 1, "app_id": APP_ID, "version": version,
                "repository": info["repository"], "python_min": [3, 11],
                "files": {name: hashlib.sha256(data).hexdigest() for name, data in payloads.items()}}
    output = output or root / "dist"
    output.mkdir(parents=True, exist_ok=True)
    archive = output / f"OrganizadorDeVista-{version}.zip"
    temporary = archive.with_suffix(".zip.tmp")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as package:
        for name, data in payloads.items():
            package.writestr(name, data)
        package.writestr(MANIFEST, json.dumps(manifest, indent=2))
    os.replace(temporary, archive)
    archive.with_suffix(".zip.sha256").write_text(f"{sha256(archive)}  {archive.name}\n", encoding="ascii")
    write_json(root / "versao.json", info)
    return archive


def github_token() -> str:
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if token:
        return token
    if shutil.which("gh"):
        result = subprocess.run(["gh", "auth", "token"], capture_output=True, text=True, timeout=15)
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    environment = dict(os.environ, GIT_TERMINAL_PROMPT="0", GCM_INTERACTIVE="Never")
    if shutil.which("git"):
        result = subprocess.run(["git", "credential", "fill"],
                                input="protocol=https\nhost=github.com\n\n",
                                capture_output=True, text=True, env=environment, timeout=20)
        credential = dict(line.split("=", 1) for line in result.stdout.splitlines() if "=" in line)
        if credential.get("password"):
            return credential["password"]
    raise UpdateError("Entre no GitHub neste computador com: git credential-manager github login --username calavort --browser")


def api(token: str, url: str, method="GET", data=None, binary=False):
    if urllib.parse.urlparse(url).hostname not in {"api.github.com", "uploads.github.com"}:
        raise UpdateError("Endereco de publicacao invalido.")
    body = data if binary else (json.dumps(data).encode("utf-8") if data is not None else None)
    request = urllib.request.Request(url, data=body, method=method, headers={
        "Authorization": "Bearer " + token, "User-Agent": "OrganizadorDeVista-Publisher",
        "Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2026-03-10",
        "Content-Type": "application/octet-stream" if binary else "application/json",
    })
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.load(response)


def publish_package(root: Path, archive: Path, create_repository=False, notes="") -> str:
    info = read_version(root)
    token = github_token()
    owner, name = info["repository"].split("/")
    base = f"https://api.github.com/repos/{owner}/{name}"
    try:
        repo = api(token, base)
    except urllib.error.HTTPError as exc:
        if exc.code != 404 or not create_repository:
            raise
        user = api(token, "https://api.github.com/user")
        if user["login"].lower() != owner.lower():
            raise UpdateError("A conta autenticada e diferente da conta configurada para publicar.")
        repo = api(token, "https://api.github.com/user/repos", "POST", {
            "name": name, "description": "Distribuicao e atualizacoes do Organizador de Vista",
            "private": False, "auto_init": True,
        })
    if repo.get("private"):
        raise UpdateError("Este atualizador usa releases publicos. Configure um repositorio publico.")
    try:
        latest = api(token, base + "/releases/latest")
        if version_tuple(info["version"]) <= version_tuple(latest["tag_name"]):
            raise UpdateError("Esta versao ja foi publicada. Incremente --versao antes de publicar.")
    except urllib.error.HTTPError as exc:
        if exc.code != 404:
            raise
    release = api(token, base + "/releases", "POST", {
        "tag_name": "v" + info["version"], "name": "Organizador de Vista " + info["version"],
        "draft": True, "prerelease": False, "target_commitish": repo["default_branch"],
        "body": "Pacote portatil do Organizador de Vista para Windows (Python 3.11+, PySide6 e Tekla Structures).\n\n"
                "Na primeira instalacao, extraia o ZIP e execute Instalar dependencias.bat. "
                "Abra por Abrir Organizador de Vista.bat. As proximas versoes sao verificadas ao abrir o programa.\n\n"
                "Relatorios, o config.json e o tekla-root.txt nao fazem parte deste pacote: "
                "sao seus e nao sao tocados pela atualizacao."
                + ("\n\n" + notes if notes else ""),
    })
    upload = release["upload_url"].split("{")[0]
    # Keep the release as a draft until both the package and checksum are uploaded.
    for asset in (archive, archive.with_suffix(".zip.sha256")):
        uploaded = api(token, upload + "?" + urllib.parse.urlencode({"name": asset.name}),
                       "POST", asset.read_bytes(), binary=True)
        expected = "sha256:" + sha256(asset)
        if uploaded.get("size") != asset.stat().st_size or (uploaded.get("digest") and uploaded["digest"] != expected):
            raise UpdateError("Upload nao conferiu. O release foi mantido como rascunho no GitHub.")
    published = api(token, base + "/releases/" + str(release["id"]), "PATCH",
                    {"draft": False, "make_latest": "true"})
    return published["html_url"]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--versao", help="Nova versao, por exemplo 7.1.1")
    parser.add_argument("--publicar", action="store_true", help="Publica o pacote no GitHub")
    parser.add_argument("--notas", type=Path, help="Arquivo Markdown com as notas da versao")
    parser.add_argument("--criar-repositorio", action="store_true", help="Cria o repositorio publico se ainda nao existir")
    args = parser.parse_args()
    try:
        archive = build_package(ROOT, args.versao)
        print("Pacote pronto:", archive)
        print("SHA-256:", sha256(archive))
        if args.publicar:
            notes = args.notas.read_text(encoding="utf-8") if args.notas else ""
            print("Release publicado:", publish_package(ROOT, archive, args.criar_repositorio, notes))
    except Exception as exc:
        parser.exit(1, f"Publicacao nao concluida: {exc}\n")


if __name__ == "__main__":
    main()
