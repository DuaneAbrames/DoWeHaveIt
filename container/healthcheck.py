"""No database calls or scans; long-running/failed scans do not affect health."""
import json
import os
import shutil

from runtime import STATE, proc_identity


def main():
    try:
        state = json.loads((STATE / 'scheduler.json').read_text())
        if proc_identity(state['pid']) != state['identity']:
            return 1
        os.kill(state['pid'], 0)
        for command in ('pwsh', 'python3', 'flock', 'dowehaveit-run'):
            if not shutil.which(command):
                return 1
        return 0
    except (OSError, ValueError, KeyError):
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
