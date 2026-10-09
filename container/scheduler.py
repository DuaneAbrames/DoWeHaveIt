"""UTC epoch-aligned scheduler. Never replays downtime or shifts after a scan."""
import argparse
import os
import signal
import subprocess
import threading
import time

from runtime import (STATE, boundary, boundary_text, configuration, log,
                     proc_identity, stop_active_manual, write_state)


class Schedule:
    def __init__(self, now_ns, frequency, offset):
        self.frequency = frequency
        self.offset = offset
        self.next = boundary(now_ns, frequency, offset, inclusive=True)
        self.last = -1

    def tick(self, now_ns):
        if self.next is None or now_ns < self.next:
            return None
        period = self.frequency * 60 * 1_000_000_000
        latest = boundary(now_ns, self.frequency, self.offset, inclusive=True)
        if latest > now_ns:
            latest -= period
        candidate = self.next
        self.next = boundary(now_ns, self.frequency, self.offset)
        # Backward adjustments cannot rerun previously considered boundaries.
        self.next = max(self.next, self.last + period)
        if latest <= self.last:
            return None
        self.last = latest
        if now_ns - latest > 2_000_000_000:
            return ('missed', candidate)
        return ('scheduled', latest)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--check-config', action='store_true')
    args = parser.parse_args()
    try:
        frequency, original, offset, _ = configuration()
    except ValueError as exc:
        parser.exit(2, f'Configuration error: {exc}\n')
    if args.check_config:
        log(f'configuration valid frequency={frequency} original_offset={original} normalized_offset={offset}')
        return 0
    stopping = threading.Event()
    for sig in (signal.SIGTERM, signal.SIGINT):
        signal.signal(sig, lambda *_: stopping.set())
    schedule = Schedule(time.time_ns(), frequency, offset)
    log(f'scheduler frequency={frequency} original_offset={original} normalized_offset={offset} '
        f'timezone={os.environ.get("TZ", "UTC")} enabled={bool(frequency)}')
    log(f'next execution={boundary_text(schedule.next)}' if frequency else 'automatic scheduling disabled; manual scans remain available')
    write_state(STATE / 'scheduler.json', {'pid': os.getpid(), 'identity': proc_identity(os.getpid())})
    children = set()
    try:
        while not stopping.is_set():
            event = schedule.tick(time.time_ns())
            if event:
                kind, timestamp = event
                log(f'{kind} execution boundary={boundary_text(timestamp)}')
                if kind == 'scheduled':
                    try:
                        children.add(subprocess.Popen(['/usr/local/bin/dowehaveit-run', 'regular', '--scheduled'], start_new_session=True))
                    except OSError as exc:
                        log(f'failed to launch wrapper: {exc}')
            for child in list(children):
                if child.poll() is not None:
                    if child.returncode not in (0, 75):
                        log(f'failed wrapper exit_status={child.returncode}; scheduler continues')
                    children.remove(child)
            # Short bounded waits follow real UTC time, including forward/backward corrections.
            stopping.wait(0.25)
    finally:
        log('shutdown requested; stopping active worker')
        active = stop_active_manual()
        for child in children:
            try:
                os.killpg(child.pid, signal.SIGTERM)
            except ProcessLookupError:
                pass
        deadline = time.monotonic() + 25
        for child in children:
            try:
                child.wait(timeout=max(0.1, deadline - time.monotonic()))
            except subprocess.TimeoutExpired:
                os.killpg(child.pid, signal.SIGKILL)
                child.wait()
        # A docker-exec wrapper is not a scheduler child. Give it time to forward
        # SIGTERM and finish before PID 1 exits and Docker kills remaining tasks.
        while active and proc_identity(active['pid']) == active['identity'] and time.monotonic() < deadline:
            time.sleep(0.1)
        (STATE / 'scheduler.json').unlink(missing_ok=True)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
