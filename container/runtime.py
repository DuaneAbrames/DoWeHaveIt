"""Small shared container utilities; no inventory or database logic."""
import datetime as dt
import json
import os
import re
import signal
import time
from pathlib import Path

STATE = Path('/run/dowehaveit')
LOGS = Path('/config/logs')
MAX_MINUTES = (2**63 - 1) // 60


def integer(name, value, maximum=MAX_MINUTES):
    # Bound digits before converting, including adversarial huge environment values.
    if not re.fullmatch(r'[0-9]{1,19}', value) or int(value) > maximum:
        raise ValueError(f'{name} must be a nonnegative integer <= {maximum}')
    return int(value)


def configuration(env=None):
    env = os.environ if env is None else env
    frequency = integer('SCHEDULE_FREQUENCY', env.get('SCHEDULE_FREQUENCY', '20'))
    original = integer('SCHEDULE_OFFSET', env.get('SCHEDULE_OFFSET', '0'))
    retention = integer('LOG_RETENTION_DAYS', env.get('LOG_RETENTION_DAYS', '30'), 36500)
    if retention == 0:
        raise ValueError('LOG_RETENTION_DAYS must be at least 1')
    return frequency, original, original % frequency if frequency else original, retention


def stamp():
    return dt.datetime.now().astimezone().isoformat(timespec='seconds')


def log(message):
    print(f'{stamp()} {message}', flush=True)


def boundary(now_ns, frequency, offset, inclusive=False):
    if frequency == 0:
        return None
    period = frequency * 60 * 1_000_000_000
    phase = (offset % frequency) * 60 * 1_000_000_000
    quotient, remainder = divmod(now_ns - phase, period)
    if remainder or not inclusive:
        quotient += 1
    # Python integers have arbitrary precision; no timestamp multiplication overflow.
    return max(phase + quotient * period, phase)


def boundary_text(ns):
    try:
        return dt.datetime.fromtimestamp(ns // 1_000_000_000, dt.timezone.utc).astimezone().isoformat()
    except (OverflowError, OSError, ValueError):
        return f'epoch_seconds={ns // 1_000_000_000}'


def proc_identity(pid):
    try:
        # The command name can contain spaces and parentheses.
        return Path(f'/proc/{pid}/stat').read_text().rsplit(')', 1)[1].split()[19]
    except (OSError, IndexError):
        return None


def write_state(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(f'.{os.getpid()}.tmp')
    temporary.write_text(json.dumps(data))
    temporary.replace(path)


def stop_active_manual():
    try:
        active = json.loads((STATE / 'active.json').read_text())
        if proc_identity(active['pid']) == active['identity']:
            os.kill(active['pid'], signal.SIGTERM)
    except (OSError, ValueError, KeyError):
        pass


def retain_logs(days, now=None, directory=None):
    directory = LOGS if directory is None else directory
    cutoff = (time.time() if now is None else now) - days * 86400
    for path in directory.glob('worker-*.log'):
        try:
            if path.is_file() and not path.is_symlink() and path.stat().st_mtime < cutoff:
                path.unlink()
        except FileNotFoundError:
            pass
