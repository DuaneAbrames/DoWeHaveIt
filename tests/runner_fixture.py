"""Runs the real wrapper with a harmless child; never touches /config or inventory."""
from pathlib import Path
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'container'))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import run

directory, mode, code, duration = sys.argv[1:]
run.LOGS = Path(directory) / 'logs'
run.STATE = Path(directory) / 'state'
actual_popen = subprocess.Popen


def harmless_child(command, **kwargs):
    # Confirm wrapper argument selection before substituting only the external worker.
    assert command[4] == mode
    assert ('-Music' in command) == (mode == 'music')
    program = 'import sys,time; print("native stdout",flush=True); print("native stderr",file=sys.stderr,flush=True); time.sleep(float(sys.argv[1])); sys.exit(int(sys.argv[2]))'
    return actual_popen([sys.executable, '-c', program, duration, code], **kwargs)


run.subprocess.Popen = harmless_child
sys.argv = ['dowehaveit-run', mode]
raise SystemExit(run.main())
