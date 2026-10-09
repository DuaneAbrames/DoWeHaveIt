import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'container'))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from runtime import retain_logs
from preflight import mounts, validate_mounts, validate_sources


MOUNTS = '''24 1 8:1 / /mnt ro,relatime master:1 - ext4 /dev/sda ro
25 1 8:1 /appdata /config rw,relatime - ext4 /dev/sda rw
26 24 0:21 / /mnt/user ro,relatime master:2 - fuse.shfs shfs rw
27 24 0:22 / /mnt/remotes/WHELP_SwitchRoms ro,relatime master:3 - cifs //whelp/SwitchRoms rw
'''


class RuntimeTests(unittest.TestCase):
    def test_nested_readonly_and_separate_config(self):
        validate_mounts(mounts(MOUNTS), '6.6.87-unraid')
        self.assertEqual(mounts(MOUNTS.replace('/mnt/user', '/mnt/a\\040b'))[2]['target'], '/mnt/a b')

    def test_unsafe_mounts_rejected(self):
        for text in (MOUNTS.replace('ro,relatime master:2', 'rw,relatime master:2'),
                     MOUNTS.replace(' master:1', ''),
                     MOUNTS.replace('/config rw,', '/config ro,'),
                     MOUNTS.replace('/mnt ro,', '/mnt rw,')):
            with self.assertRaises(RuntimeError):
                validate_mounts(mounts(text), '6.6.87')
        with self.assertRaises(RuntimeError):
            validate_mounts(mounts(MOUNTS), '5.10.0')

    def test_missing_switch_mount_fails_before_scan(self):
        with self.assertRaisesRegex(RuntimeError, 'Remote Switch mount'):
            validate_sources('regular', mounts('\n'.join(MOUNTS.splitlines()[:-1])))

    def test_log_retention(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            old = root / 'worker-old.log'
            current = root / 'worker-current.log'
            unrelated = root / 'notes.log'
            for path in (old, current, unrelated):
                path.write_text('data')
            os.utime(old, (1, 1))
            retain_logs(30, now=3_000_000, directory=root)
            self.assertFalse(old.exists())
            self.assertTrue(current.exists())
            self.assertTrue(unrelated.exists())


@unittest.skipUnless(sys.platform == 'linux', 'flock and process-group tests require Linux')
class LinuxExecutionTests(unittest.TestCase):
    def test_password_redaction_across_output_chunks(self):
        from run import Redactor
        password = b's3cret-password'
        original = b'connection=' + password + b'; error ' + password + b'\n'
        for length in range(1, len(original) + 1):
            redactor = Redactor([password])
            result = b''.join(redactor.feed(original[i:i+length]) for i in range(0, len(original), length))
            result += redactor.feed(b'', final=True)
            self.assertNotIn(password, result)
            self.assertEqual(result, original.replace(password, b'[REDACTED]'))

    def test_lock_is_shared_and_released(self):
        from run import acquire_lock
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'worker.lock'
            first = acquire_lock(path)
            self.assertIsNotNone(first)
            self.assertIsNone(acquire_lock(path))
            first.close()
            next_lock = acquire_lock(path)
            self.assertIsNotNone(next_lock)
            next_lock.close()

    def fixture(self, directory, mode='regular', code=0, duration=0):
        return subprocess.Popen([sys.executable, str(Path(__file__).with_name('runner_fixture.py')),
                                 directory, mode, str(code), str(duration)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

    def test_worker_status_output_mode_and_failures(self):
        for mode, code in (('regular', 0), ('music', 7)):
            with tempfile.TemporaryDirectory() as directory:
                child = self.fixture(directory, mode, code)
                output, _ = child.communicate(timeout=10)
                self.assertEqual(child.returncode, code)
                text = output.decode()
                self.assertIn(f'mode={mode}', text)
                self.assertIn('native stdout', text)
                self.assertIn('native stderr', text)
                self.assertIn(f'exit_status={code}', text)
                logs = list((Path(directory) / 'logs').glob('worker-*.log'))
                self.assertEqual(len(logs), 1)
                self.assertEqual(logs[0].read_bytes(), output)

    def test_overlapping_workers_and_shutdown(self):
        with tempfile.TemporaryDirectory() as directory:
            first = self.fixture(directory, duration=60)
            try:
                active = Path(directory) / 'state/active.json'
                deadline = time.monotonic() + 5
                while not active.exists() and time.monotonic() < deadline:
                    time.sleep(0.01)
                self.assertTrue(active.exists())
                second = self.fixture(directory, 'music')
                output, _ = second.communicate(timeout=5)
                self.assertEqual(second.returncode, 75)
                self.assertIn(b'skipped', output)
                first.terminate()
                output, _ = first.communicate(timeout=20)
                self.assertEqual(first.returncode, 143)
                self.assertIn(b'exit_status=143', output)
                self.assertFalse(active.exists())
                third = self.fixture(directory)
                third.communicate(timeout=5)
                self.assertEqual(third.returncode, 0)
            finally:
                if first.poll() is None:
                    first.kill()
                first.communicate()

    def test_disabled_scheduler_health_and_manual_shutdown(self):
        module_dir = str(Path(sys.modules['runtime'].__file__).resolve().parent)
        with tempfile.TemporaryDirectory() as directory:
            manual = self.fixture(directory, duration=60)
            active = Path(directory) / 'state/active.json'
            deadline = time.monotonic() + 5
            while not active.exists() and time.monotonic() < deadline:
                time.sleep(0.01)
            self.assertTrue(active.exists())
            program = ('import os,sys; from pathlib import Path; '
                       'sys.path.insert(0,sys.argv[1]); import runtime,scheduler; '
                       'runtime.STATE=Path(sys.argv[2]); scheduler.STATE=runtime.STATE; '
                       'os.environ["SCHEDULE_FREQUENCY"]="0"; sys.argv=["scheduler"]; scheduler.main()')
            scheduler = subprocess.Popen([sys.executable, '-c', program, module_dir, str(active.parent)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            try:
                state = active.parent / 'scheduler.json'
                deadline = time.monotonic() + 5
                while not state.exists() and time.monotonic() < deadline:
                    time.sleep(0.01)
                self.assertTrue(state.exists())
                import healthcheck
                with patch.object(healthcheck, 'STATE', active.parent), patch.object(healthcheck.shutil, 'which', return_value='/bin/true'):
                    self.assertEqual(healthcheck.main(), 0)
                scheduler.terminate()
                manual.communicate(timeout=20)
                output, _ = scheduler.communicate(timeout=30)
                self.assertEqual(scheduler.returncode, 0)
                self.assertIn(b'automatic scheduling disabled', output)
                self.assertEqual(manual.returncode, 143)
                self.assertFalse(state.exists())
                self.assertFalse(active.exists())
            finally:
                for process in (scheduler, manual):
                    if process.poll() is None:
                        process.kill()
                    process.communicate()


if __name__ == '__main__':
    unittest.main()
