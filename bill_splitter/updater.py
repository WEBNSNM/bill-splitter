"""Opt-in GitHub Releases updates; no bill data leaves the computer."""
import hashlib
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from urllib.parse import quote

from .version import VERSION, GITHUB_REPOSITORY

EXE_NAME = 'wuli娟娟.exe'
CHECKSUM_NAME = EXE_NAME + '.sha256'


def version_tuple(value):
    if not re.fullmatch(r'v?\d+\.\d+\.\d+', value):
        raise ValueError('版本号必须为 x.y.z')
    return tuple(map(int, value.removeprefix('v').split('.')))


def select_release(release, current=VERSION):
    if release.get('draft') or release.get('prerelease'):
        return None
    if version_tuple(release['tag_name']) <= version_tuple(current):
        return None
    assets = {a['name']: a['browser_download_url'] for a in release['assets']}
    if not all(n in assets for n in (EXE_NAME, CHECKSUM_NAME)):
        raise ValueError('新版本尚未提供完整的 EXE 和校验文件，请稍后重试。')
    for url in assets.values():
        if not url.startswith(f'https://github.com/{GITHUB_REPOSITORY}/releases/download/'):
            raise ValueError('更新下载地址不属于本项目。')
    return {'version': release['tag_name'], 'assets': assets}


def fetch(url):
    request = Request(url, headers={'User-Agent': 'BillSplitter/' + VERSION})
    return urlopen(request, timeout=30)


def check_update():
    try:
        with fetch(f'https://github.com/{GITHUB_REPOSITORY}/releases.atom') as response:
            data = response.read()
            root = ET.fromstring(data)
            if not root.findall('{http://www.w3.org/2005/Atom}entry'):
                raise ValueError('尚无公开发布版本，暂时无法提供在线更新；请等待项目发布正式版本。')
            return parse_release_feed(data)
    except HTTPError as error:
        if error.code == 404:
            raise ValueError('尚无公开正式版本，或仓库尚未公开。') from error
        raise


def release_from_tag(tag, link):
    expected = f'https://github.com/{GITHUB_REPOSITORY}/releases/tag/'
    if not link.startswith(expected):
        raise ValueError('GitHub 更新订阅源地址无效。')
    if version_tuple(tag) <= version_tuple(VERSION):
        return None
    base = f'https://github.com/{GITHUB_REPOSITORY}/releases/download/{quote(tag, safe="")}'
    assets = {name: f'{base}/{quote(name, safe="")}' for name in (EXE_NAME, CHECKSUM_NAME)}
    return {'version': tag, 'assets': assets}


def parse_release_feed(data):
    root = ET.fromstring(data)
    namespace = {'atom': 'http://www.w3.org/2005/Atom'}
    entries = root.findall('atom:entry', namespace)
    if not entries:
        return None
    entry = entries[0]
    tag = (entry.findtext('atom:title', default='', namespaces=namespace)).strip()
    link = entry.find('atom:link', namespace)
    if link is None or not tag:
        raise ValueError('无法读取 GitHub 最新版本信息。')
    return release_from_tag(tag, link.attrib.get('href', ''))


def verify_checksum(path, checksum):
    expected = checksum.strip().split()[0] if checksum.strip() else ''
    if not re.fullmatch('[0-9a-fA-F]{64}', expected):
        raise ValueError('更新校验文件无效。')
    with path.open('rb') as source:
        digest = hashlib.file_digest(source, 'sha256')
    if digest.hexdigest().lower() != expected.lower():
        raise ValueError('更新文件校验失败，已取消安装，请重新检查更新。')


def download_update(release):
    folder = Path(tempfile.mkdtemp(prefix='BillSplitter-update-'))
    target = folder / EXE_NAME
    with fetch(release['assets'][CHECKSUM_NAME]) as response:
        checksum = response.read(4096).decode('ascii')
    with fetch(release['assets'][EXE_NAME]) as response, target.open('wb') as output:
        while block := response.read(1024 * 1024):
            output.write(block)
    verify_checksum(target, checksum)
    return target


def install_update(download):
    if not getattr(sys, 'frozen', False) or sys.platform != 'win32':
        raise ValueError('源码模式请使用下载的 EXE；自动替换仅支持 Windows EXE。')
    helper = download.parent / 'install.ps1'
    helper.write_text(r'''param([int]$AppPid, [string]$Target, [string]$Download)
$ErrorActionPreference = 'Stop'
Wait-Process -Id $AppPid -ErrorAction SilentlyContinue
$backup = $Target + '.previous'
$installed = $false
try {
    Copy-Item -LiteralPath $Target -Destination $backup -Force
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        try {
            Copy-Item -LiteralPath $Download -Destination $Target -Force
            $installed = $true
            break
        } catch { Start-Sleep -Milliseconds 1000 }
    }
    if (-not $installed) { throw '无法替换程序，请检查文件夹写入权限。' }
    Start-Process -FilePath $Target
} catch {
    if (Test-Path -LiteralPath $backup) { Copy-Item -LiteralPath $backup -Destination $Target -Force -ErrorAction SilentlyContinue }
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show('更新失败，旧版备份位于：' + $backup + "`n" + $_.Exception.Message, 'wuli娟娟更新')
}
''', encoding='utf-8-sig')
    import os
    subprocess.Popen(['powershell.exe', '-NoProfile', '-WindowStyle', 'Hidden',
                      '-ExecutionPolicy', 'Bypass', '-File', str(helper),
                      '-AppPid', str(os.getpid()), '-Target', sys.executable,
                      '-Download', str(download)], creationflags=subprocess.CREATE_NO_WINDOW)
