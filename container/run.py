"""One execution: common lock, source validation, priorities, logs and status."""
import argparse
import fcntl
import os
import re
import signal
import subprocess
import sys
import threading
import time
import uuid

from runtime import LOGS, STATE, configuration, proc_identity, retain_logs, stamp, write_state


class Redactor:
    def __init__(self, secrets):
        self.secrets = [value for value in secrets if value]
        self.keep = max((len(value) for value in self.secrets), default=1) - 1
        self.pending = b''

    def feed(self, data, final=False):
        self.pending += data
        for value in self.secrets:
            self.pending = self.pending.replace(value, b'[REDACTED]')
        count = len(self.pending) if final else max(0, len(self.pending) - self.keep)
        result, self.pending = self.pending[:count], self.pending[count:]
        return result


def passwords():
    from pathlib import Path
    config = Path('/config/database-info.txt')
    if not config.is_file():
        config = Path('/mnt/user/bigone/scripts/DoWeHaveIt/database-info.txt')
    try:
        result = []
        for line in config.read_text(encoding='utf-8-sig').splitlines():
            match = re.match(r'^\s*db_pass:\s*(.*)$', line, re.IGNORECASE)
            if match:
                result.append(match[1].strip().encode())
        return result
    except OSError:
        return []


def acquire_lock(path):
    stream = path.open('a+')
    try:
        fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        return stream
    except BlockingIOError:
        stream.close()
        return None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('mode', nargs='?', choices=('regular', 'music'), default='regular')
    parser.add_argument('--scheduled', action='store_true')
    args = parser.parse_args()
    _, _, _, retention = configuration()
    LOGS.mkdir(parents=True, exist_ok=True)
    STATE.mkdir(parents=True, exist_ok=True)
    log_path = LOGS / f'worker-{time.time_ns()}-{args.mode}-{uuid.uuid4().hex[:8]}.log'
    started = time.monotonic()
    started_at = stamp()
    stopping = threading.Event()
    for sig in (signal.SIGTERM, signal.SIGINT):
        signal.signal(sig, lambda *_: stopping.set())
    with log_path.open('xb') as output:
        def emit(data):
            output.write(data)
            output.flush()
            sys.stdout.buffer.write(data)
            sys.stdout.buffer.flush()

        def message(text):
            emit(f'{stamp()} {text}\n'.encode())

        retain_logs(retention, directory=LOGS)
        lock = acquire_lock(STATE / 'worker.lock')
        if lock is None:
            message(f'skipped mode={args.mode} source={"scheduled" if args.scheduled else "manual"} reason=active-lock started_at={started_at} ended_at={stamp()} exit_status=75 duration_seconds={time.monotonic()-started:.3f}')
            return 75
        with lock:
            identity = proc_identity(os.getpid())
            write_state(STATE / 'active.json', {'pid': os.getpid(), 'identity': identity, 'mode': args.mode, 'log': str(log_path)})
            message(f'started mode={args.mode} source={"scheduled" if args.scheduled else "manual"} log={log_path.name}')
            status = 1
            child = None
            try:
                command = ['bash', '-c',
                           'python3 /opt/dowehaveit/preflight.py --mode "$1" && '
                           'exec nice -n 10 ionice -c 2 -n 7 pwsh -NoLogo -NoProfile -File /opt/dowehaveit/invoke-worker.ps1 "${@:2}"',
                           'dowehaveit-worker', args.mode]
                if args.mode == 'music':
                    command.append('-Music')
                if stopping.is_set():
                    status = 143
                else:
                    child = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                             start_new_session=True, pass_fds=(lock.fileno(),))
                    # Stream binary chunks without line buffering or losing native stderr.
                    def copy_output():
                        redactor = Redactor(passwords())
                        while chunk := child.stdout.read1(65536):
                            emit(redactor.feed(chunk))
                        emit(redactor.feed(b'', final=True))

                    reader = threading.Thread(target=copy_output, daemon=True)
                    reader.start()
                    while child.poll() is None:
                        if stopping.wait(0.1):
                            os.killpg(child.pid, signal.SIGTERM)
                            try:
                                child.wait(timeout=15)
                            except subprocess.TimeoutExpired:
                                os.killpg(child.pid, signal.SIGKILL)
                            break
                    code = child.wait()
                    status = 143 if stopping.is_set() else (128 - code if code < 0 else code)
                    reader.join(timeout=2)
                    child.stdout.close()
            except OSError as exc:
                message(f'failed to execute worker: {exc}')
            finally:
                message(f'{"completed" if status == 0 else "failed"} mode={args.mode} exit_status={status} duration_seconds={time.monotonic()-started:.3f}')
                write_state(STATE / 'last-run.json', {'mode': args.mode, 'exit_status': status, 'ended': stamp(), 'log': str(log_path)})
                (STATE / 'active.json').unlink(missing_ok=True)
            return status


if __name__ == '__main__':
    raise SystemExit(main())
