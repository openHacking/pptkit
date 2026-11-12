#!/usr/bin/env python3
"""Run the native corpus audit's soffice calls in an isolated Debian container."""
import os
import sys
from pathlib import Path
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[2]
IMAGE = os.environ.get('PPTKIT_LIBREOFFICE_IMAGE', 'pptkit-lo-trixie-arm64:compat')


def main(args):
    command = ['docker', 'run', '--rm', '--pull=never', '--network=none']
    if args == ['--version']:
        os.execvp('docker', [*command, IMAGE, *args])
    option = '-env:UserInstallation='
    if not args or not args[0].startswith(option):
        raise SystemExit('Expected a native corpus audit profile')
    uri = urlsplit(args[0][len(option):])
    profile = Path(unquote(uri.path)).resolve()
    work = profile.parent
    if uri.scheme != 'file' or uri.netloc or work.parent != ROOT / '.tmp' or not work.name.startswith('corpus-native-resave-'):
        raise SystemExit('Expected a profile in the current corpus audit directory')
    sources = ROOT / '.tmp/fixtures'
    command += ['--user', f'{os.getuid()}:{os.getgid()}', '-e', 'HOME=/tmp',
                '--mount', f'type=bind,src={sources},dst={sources},readonly',
                '--mount', f'type=bind,src={work},dst={work}']
    os.execvp('docker', [*command, IMAGE, *args])


if __name__ == '__main__':
    main(sys.argv[1:])
