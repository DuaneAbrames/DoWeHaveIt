"""Validate mount safety and source availability without scanning inventories."""
import argparse
import os
import re
import subprocess
import tempfile
from pathlib import Path

from runtime import LOGS, STATE, log


def unescape(value):
    return re.sub(r'\\([0-7]{3})', lambda match: chr(int(match[1], 8)), value)


def mounts(text):
    result = []
    for line in text.splitlines():
        fields = line.split()
        separator = fields.index('-')
        result.append({'target': unescape(fields[4]), 'options': fields[5].split(','),
                       'optional': fields[6:separator], 'type': fields[separator + 1]})
    return result


def validate_mounts(entries, kernel):
    version = re.match(r'(\d+)\.(\d+)', kernel)
    if not version or tuple(map(int, version.groups())) < (5, 12):
        raise RuntimeError('Linux kernel >= 5.12 is required for recursive read-only binds; upgrade the host')
    media = [m for m in entries if m['target'] == '/mnt' or m['target'].startswith('/mnt/')]
    root = [m for m in media if m['target'] == '/mnt']
    if len(root) != 1 or 'ro' not in root[0]['options']:
        raise RuntimeError('/mnt must be a separate read-only bind mount')
    if not any(flag.startswith('master:') for flag in root[0]['optional']):
        raise RuntimeError('/mnt must use rslave propagation from a shared/slave host mount; inspect host propagation')
    writable = [m['target'] for m in media if 'ro' not in m['options']]
    if writable:
        raise RuntimeError('Writable child mounts under /mnt: ' + ', '.join(writable))
    config = [m for m in entries if m['target'] == '/config']
    if len(config) != 1 or 'rw' not in config[0]['options']:
        raise RuntimeError('/config must be a separate writable bind mount')


def validate_sources(mode, entries):
    required = ['/mnt/user/media/mp3']
    if mode == 'regular':
        required += ['/mnt/user/media/Movies', '/mnt/user/media/TV Shows',
                     '/mnt/user/bigone/DISC LIST', '/mnt/user/bigone/Bittorrent/Completed/GoG',
                     '/mnt/user/bigone/Bittorrent/Completed/SteamUnlocked',
                     '/mnt/remotes/WHELP_SwitchRoms/Base', '/mnt/remotes/WHELP_SwitchRoms/Update',
                     '/mnt/remotes/WHELP_SwitchRoms/DLC']
    targets = {m['target'] for m in entries}
    if '/mnt/user' not in targets:
        raise RuntimeError('Unraid /mnt/user mount is unavailable')
    if mode == 'regular' and '/mnt/remotes/WHELP_SwitchRoms' not in targets:
        raise RuntimeError('Remote Switch mount is unavailable; refusing an incomplete regular inventory')
    for value in required:
        path = Path(value)
        if not path.is_dir():
            raise RuntimeError(f'Required source directory unavailable: {path}')
        # Read one entry to check accessibility; do not enumerate the tree.
        with os.scandir(path) as items:
            next(items, None)
    if mode == 'regular':
        for name in ('disc list.csv', 'Collection.xml'):
            with (Path('/mnt/user/bigone/DISC LIST') / name).open('rb') as stream:
                stream.read(1)


def check(startup=False, mode='regular'):
    entries = mounts(Path('/proc/self/mountinfo').read_text())
    validate_mounts(entries, os.uname().release)
    for name in ('DoWeHaveIt-Worker.ps1', 'DoWeHaveIt-Common.ps1'):
        if not (Path('/config') / name).is_file():
            raise RuntimeError(f'Missing /config/{name}')
    config = Path('/config/database-info.txt')
    legacy = Path('/mnt/user/bigone/scripts/DoWeHaveIt/database-info.txt')
    if not config.is_file() and not legacy.is_file():
        raise RuntimeError('Database configuration missing from /config and legacy script directory')
    LOGS.mkdir(parents=True, exist_ok=True)
    STATE.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryFile(dir=LOGS) as stream:
        stream.write(b'writable')
    if not startup:
        validate_sources(mode, entries)
    subprocess.run(['pwsh', '-NoLogo', '-NoProfile', '-Command',
                    '$ErrorActionPreference="Stop"; Import-Module SimplySql -ErrorAction Stop; Get-Command Open-MySqlConnection -ErrorAction Stop | Out-Null'], check=True)
    log(f'preflight passed mode={mode} startup={startup}')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--startup', action='store_true')
    parser.add_argument('--mode', choices=('regular', 'music'), default='regular')
    args = parser.parse_args()
    try:
        check(args.startup, args.mode)
        return 0
    except (OSError, RuntimeError, subprocess.CalledProcessError) as exc:
        log(f'preflight failed: {exc}')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
