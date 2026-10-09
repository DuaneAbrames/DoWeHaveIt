# DoWeHaveIt container deployment

This image packages the Linux AMD64 runtime, SimplySql, diagnostics, and scheduling support. Application scripts and database configuration stay on the host at `/config`. No inbound ports, graphical components, privileged mode, extra capabilities, or production build-time scans are required.

## Runtime and image publication

The Dockerfile uses `debian:12-slim`, Microsoft’s signed Debian package repository, PowerShell `7.6.6-1.deb`, and SimplySql `2.2.0.106`. SimplySql is installed during the build in the shared AllUsers module location, outside `/config` and user profiles. The build verifies PowerShell, module import, `Open-MySqlConnection`, diagnostic executables, wrapper permissions, scheduler configuration, tests, and bootstrap parsing. Version updates are explicit changes to the Dockerfile followed by validation.

The image contains no application scripts, CSV/XML datasets, or database credentials. `.dockerignore` permits only container support and tests into the build context. Tests bundled with the image are useful for diagnostics and never execute the inventory application.

The [Microsoft Debian installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/install-debian?view=powershell-7.6) documents supported installation and shared module paths. The pinned module is listed on [PowerShell Gallery](https://www.powershellgallery.com/packages/SimplySql/2.2.0.106).

GitHub Actions derives the image as `ghcr.io/<lowercase-repository-owner>/<lowercase-repository-name>`. For this repository it is:

```text
ghcr.io/duaneabrames/dowehaveit:latest
```

Every push and pull request validates on a GitHub-hosted Linux runner. Only validated default-branch pushes/manual dispatches publish `latest`. Publishing a GitHub release produces its release tag (for example `v2.0.0`) and its semantic version tag (`2.0.0` when applicable); releases do not replace `latest`. SHA tags and OCI source/revision labels identify builds. Buildx uses GitHub Actions caching. Only the publication job receives `packages: write`; publication uses `GITHUB_TOKEN`, following [GitHub’s GHCR publication guidance](https://docs.github.com/en/actions/tutorials/publish-packages/publish-docker-images). Pull requests never publish.

After the first successful workflow, open the repository’s linked package in GitHub. For public pulls, change the package visibility to **Public** in package settings, then:

```bash
docker pull ghcr.io/duaneabrames/dowehaveit:latest
```

For a private package, grant the administrator’s GitHub account package read access. On Unraid, authenticate with a classic PAT having `read:packages`, authorizing organization SSO if required. This pull credential is separate from Actions publication, which needs no PAT. Never put it in Compose, templates, image arguments, or application files:

```bash
read -rsp 'GHCR read token: ' GHCR_READ_TOKEN; echo
printf '%s' "$GHCR_READ_TOKEN" | docker login ghcr.io -u YOUR_GITHUB_USERNAME --password-stdin
unset GHCR_READ_TOKEN
docker pull ghcr.io/duaneabrames/dowehaveit:latest
```

See [GitHub container registry authentication and visibility](https://docs.github.com/en/packages/working-with-a-github-packages-registry/working-with-the-container-registry). Pin a release or image digest for controlled production upgrades. Runtime source files are independently supplied through `/config`; deploy the compatible Stage 1 files with the small Stage 2 Common change below.

## Prepare Unraid and validate mount protection

Use Linux kernel **5.12 or newer** and a current Docker Engine supporting recursive read-only binds (Docker 25 or newer). On the host:

```bash
uname -r
docker version
findmnt -o TARGET,FSTYPE,OPTIONS,PROPAGATION /mnt
findmnt -o TARGET,FSTYPE,OPTIONS,PROPAGATION /mnt/user
findmnt -o TARGET,FSTYPE,OPTIONS,PROPAGATION /mnt/remotes/WHELP_SwitchRoms
```

The source mount containing `/mnt` must support receiving propagation from a shared/slave host mount. `rslave` passes host mount changes into the container. It cannot make an unshared host mount propagate. If preflight reports missing propagation, configure the host’s mount propagation deliberately before deployment; do not give the container mount privileges.

`/mnt:/mnt:ro,rslave` includes nested mounts, but their read-only protection depends on the kernel and Docker version. [Docker’s bind mount documentation](https://docs.docker.com/engine/storage/bind-mounts/) describes the kernel requirement and the explicit `bind-recursive=readonly` option. This implementation rejects an old kernel, a writable `/mnt`, or **any writable descendant mount** visible under `/mnt`. It checks mount-level `ro` flags from `/proc/self/mountinfo`, even if a backing filesystem’s superblock remains writable on the host. It also checks slave propagation and a separate writable `/config` mount. Unsupported configurations fail explicitly; there is no unsafe bypass.

For a stricter Docker CLI declaration, use `--mount` instead of `-v`:

```text
--mount type=bind,source=/mnt,target=/mnt,readonly,bind-propagation=rslave,bind-recursive=readonly
```

Docker rejects this option on unsupported hosts. The Compose example uses read-only bind plus `rslave` and relies on startup/per-execution mount validation to enforce child protection. Newly propagated host mounts can be writable despite existing protection: checks before every scan reject them. Recheck after mounting/remounting remote sources and recreate the container if necessary. Keep host mounts stable during a scan; container mount propagation is not an atomic snapshot of host topology.

Do not test read-only behavior by attempting to write to media. Inspect the effective mount flags instead. `/config` is a distinct writable container path; its host directory can safely be under `/mnt/user/appdata` while the separate `/mnt` bind remains read-only.

## Prepare application configuration

Create an appdata directory, for example `/mnt/user/appdata/dowehaveit`, and copy these existing application files into it:

```text
DoWeHaveIt.ps1
DoWeHaveIt-UI.ps1
DoWeHaveIt-Worker.ps1
DoWeHaveIt-Common.ps1
database-info.txt
```

Use the existing database configuration format: `host`, `port`, `database`, `db_user`, and `db_pass` keys separated from values by a colon. Secure the file on the host. Do not print its contents while debugging. Set `host` to a reachable database hostname/address; `localhost` identifies the container, not Unraid or another database container. The default bridge network has outbound access; do not add inbound ports or host networking.

Linux configuration precedence is now:

1. `database-info.txt` alongside Common (`/config/database-info.txt` in the container), when it exists.
2. The legacy `/mnt/user/bigone/scripts/DoWeHaveIt/database-info.txt` otherwise.

An invalid existing `/config` file is an error; it does not silently fall back. Windows configuration is unchanged. All inventory source paths and legacy bootstrap CSV locations remain unchanged, including `/mnt/user/media`, `/mnt/user/bigone/DISC LIST`, the BitTorrent directories, and `/mnt/remotes/WHELP_SwitchRoms`. Existing bootstrap CSVs remain optional at the legacy script location. This stage does not redirect them to appdata.

## Unraid Docker UI configuration

Add a container through Unraid’s Docker interface (enable Advanced View for extra parameters):

| Setting | Value |
| --- | --- |
| Name | `dowehaveit` |
| Repository | `ghcr.io/duaneabrames/dowehaveit:latest` or a release/digest |
| Network Type | Bridge |
| Port mappings | None |
| `/config` path | Configurable host appdata directory → `/config`, Read/Write |
| `/mnt` path | `/mnt` → `/mnt`, Read Only; propagation **rslave** |
| `TZ` variable | `America/New_York` or your timezone |
| `SCHEDULE_FREQUENCY` variable | `0` initially; `20` after verification |
| `SCHEDULE_OFFSET` variable | `0` |
| `LOG_RETENTION_DAYS` variable | `30` |
| Restart policy | `unless-stopped` |
| CPU allocation | Approximately two logical CPUs |

If the UI cannot preserve read-only/rslave options, omit its `/mnt` path mapping and put exactly one mount in Extra Parameters using the explicit `--mount` declaration above. Do not add a second `/mnt` mapping. A UI mapping capable of representing `/mnt:/mnt:ro,rslave` is also supported, subject to preflight verification. No optional XML template is supplied.

Extra Parameters for a CPU quota, orderly shutdown, and bounded Docker logs:

```text
--cpus=2 --restart=unless-stopped --stop-timeout=40 --log-opt max-size=10m --log-opt max-file=3
```

Configure each setting only once if the UI already exposes it. For affinity, use Unraid CPU pinning or `--cpuset-cpus=<host-selected-logical-CPU-list>`. CPU quota limits total CPU time; affinity selects eligible logical CPUs. They can be combined. Do not copy another host’s CPU numbers. The wrapper uses `nice -n 10` and `ionice -c 2 -n 7`, inherited by worker child processes. CPU limits and low process priority do not guarantee disk I/O isolation, and some I/O schedulers ignore `ionice`. No memory limit is imposed before measuring usage.

## Compose or CLI deployment

The supplied `compose.yaml` defaults to scheduling disabled. Set the configurable image and appdata path in your shell or a local, untracked `.env` file:

```bash
export DOWEHAVEIT_IMAGE=ghcr.io/duaneabrames/dowehaveit:latest
export DOWEHAVEIT_CONFIG=/mnt/user/appdata/dowehaveit
export SCHEDULE_FREQUENCY=0
docker compose config --quiet
docker compose up -d
```

Alternatively, on a validated host:

```bash
docker run -d --name dowehaveit --restart unless-stopped \
  --cpus=2 --stop-timeout=40 \
  --log-opt max-size=10m --log-opt max-file=3 \
  --mount type=bind,source=/mnt/user/appdata/dowehaveit,target=/config \
  --mount type=bind,source=/mnt,target=/mnt,readonly,bind-propagation=rslave,bind-recursive=readonly \
  -e TZ=America/New_York -e SCHEDULE_FREQUENCY=0 -e SCHEDULE_OFFSET=0 \
  -e LOG_RETENTION_DAYS=30 ghcr.io/duaneabrames/dowehaveit:latest
```

Mount the complete existing sources at `/mnt`; do not substitute physical array/cache paths. The application continues using its existing `Get-ChildItem` scanning logic.

## Scheduling semantics

The default image schedule is 20 minutes, offset zero. Environment values are read at startup. Recreate the container after changing its environment; restarting an existing container only rereads its existing values.

```text
scheduled_time = 1970-01-01 00:00 UTC + offset_minutes + N × frequency_minutes
```

Both schedule values must be nonnegative ASCII integers. Positive frequencies normalize offset with `offset % frequency`. Values are bounded to `153722867280912930` minutes; Python’s arbitrary-precision arithmetic prevents timestamp overflow. Log retention is an integer from 1 through 36500 days. Invalid values cause a clear startup error.

| Frequency | Offset | Schedule |
| --- | --- | --- |
| 20 | 0 | :00, :20, :40 |
| 20 | 5 | :05, :25, :45 |
| 15 | 7 | :07, :22, :37, :52 |
| 60 | 10 | :10 each hour |
| 137 | 12 | Every 137 minutes, phased from the epoch; crosses hours/days |
| 0 | Any valid value | No automatic scans; container and health check remain running |

Startup between boundaries waits for the next one. Startup exactly on a boundary may run it. No downtime replay occurs. The scheduler checks UTC wall time in short bounded waits, so completion duration does not shift execution times. Local DST affects log presentation only. A high-water boundary prevents duplicates after backward clock adjustments. Large forward jumps or suspension skip stale boundaries rather than replaying them; an execution within two seconds of the current boundary is treated as timely. Clock adjustments during a running scan do not affect its monotonic duration measurement.

Scheduled and manual wrappers share `/run/dowehaveit/worker.lock`. An active scan causes later executions to skip with exit status **75**, recorded in Docker and persistent logs. Failed workers do not terminate the scheduler. There is no cron-expression variable or scheduling framework.

Startup logs frequency, original/normalized offset, timezone, scheduling enablement, and the next boundary. Set frequency to `20` only after deployment checks, then recreate/apply the Unraid container configuration.

## Logging, health and shutdown

`docker logs dowehaveit` shows scheduler events and worker output. Every wrapper attempt creates `/config/logs/worker-<epoch-nanoseconds>-<mode>-<id>.log` containing timestamps, mode, start/completion or failure, duration, status, and merged native/PowerShell stdout and stderr. Skipped attempts record their timestamps, status and reason. The regular worker can additionally publish a music snapshot according to its unchanged 24-hour decision; the wrapper’s `regular` label identifies the requested execution mode.

Known `db_pass` values are masked in captured worker output, including across stream chunk boundaries. Do not put credentials in command arguments or print the configuration manually. Newly generated runtime logs use restrictive permissions. Retention removes matching worker logs older than `LOG_RETENTION_DAYS` on each execution attempt. Docker log rotation is configured separately in Compose/CLI/UI; it is not built into the image.

The inexpensive health check verifies the scheduler PID and process identity plus essential executable availability. It makes no database calls or inventory scans. Disabled scheduling, a long scan, and a failed scan do not by themselves make the scheduler unhealthy. Health describes runtime liveness, not database or media readiness. Use preflight for readiness checks.

On SIGTERM/SIGINT the scheduler stops launching executions and signals the active wrapper (including a manually launched wrapper). The wrapper forwards termination to its worker process group, waits up to 15 seconds, and then kills the group if necessary. The scheduler waits up to 25 seconds for its wrappers; use Docker’s 40-second stop timeout. Lock ownership is inherited by the worker until it exits. Forced termination may interrupt the application before its `finally`; server-side connection loss rolls back an open transaction, and already committed complete snapshots remain committed. Do not assume a killed run reaches the application’s normal run-status update.

## Manual execution and diagnostics

Use the supported wrapper for both manual modes; it manages locking, preflight, logging, priority and exit status:

```bash
docker exec dowehaveit dowehaveit-run regular
docker exec dowehaveit dowehaveit-run music
docker exec -it dowehaveit bash
```

Inside the container:

```bash
# Readiness checks, with no inventory or database publication.
dowehaveit-preflight --mode regular
dowehaveit-preflight --mode music

# PowerShell/module diagnostics.
pwsh -NoLogo -NoProfile -Command '$PSVersionTable; Get-Module -ListAvailable SimplySql; Import-Module SimplySql -ErrorAction Stop; Get-Command Open-MySqlConnection'

# Worker and scheduler processes, active lock metadata.
ps -eo pid,ppid,ni,etime,args | grep -E 'scheduler|run.py|pwsh'
lslocks
cat /run/dowehaveit/active.json
flock -n /run/dowehaveit/worker.lock true; echo "lock probe status=$?"

# Mount protection and propagation; inspect flags, do not write to media.
findmnt -R /mnt -o TARGET,FSTYPE,VFS-OPTIONS,PROPAGATION
findmnt /config -o TARGET,FSTYPE,VFS-OPTIONS,PROPAGATION
mountpoint /mnt/user
mountpoint /mnt/remotes/WHELP_SwitchRoms

# Network diagnostics; substitute the configured host and port.
getent hosts DB_HOST
nc -vz DB_HOST 3306
mariadb --host=DB_HOST --port=3306 --user=DB_USER --password --database=DB_NAME --execute='SELECT 1;'

# Persistent execution history and system diagnostics.
ls -lt /config/logs
tail -n 100 /config/logs/worker-*.log
cat /run/dowehaveit/last-run.json
iostat -xz 1 3
free -h
```

`mariadb --password` prompts interactively without placing the password in process arguments. For application-level connection verification, use the existing Common functions with cleanup and no initialization/publication:

```bash
pwsh -NoLogo -NoProfile -Command '$ErrorActionPreference="Stop"; . /config/DoWeHaveIt-Common.ps1; try { Open-AppConnection; Invoke-AppScalarValue "SELECT 1;" } finally { Close-AppConnection }'
```

Confirm the actual database `SELECT 1` manually in the target environment. Avoid logging database credentials or configuration content. Diagnostics such as `iostat` may show only host-visible devices; `strace` may be restricted by container security. Do not add privileged mode or kernel capabilities by default.

Raw `pwsh -NoLogo -NoProfile -File /config/DoWeHaveIt-Worker.ps1` remains possible for debugging, but **bypasses the container lock, preflight, failure policy, priorities, and persistent logging**. It is not the supported production execution path. No override switch is provided; use the wrapper to preserve overlap prevention.

## Failure handling and troubleshooting

| Failure | Behavior / response |
| --- | --- |
| Missing `/config`, scripts or both database-config locations | Startup fails explicitly; correct bind/files before restarting |
| Writable media child mounts, unsupported kernel or propagation | Startup fails; correct host/Docker/mount options, then recreate |
| Required media or remote Switch mount missing | Pre-execution checks fail; no worker starts and scheduler remains available |
| Required CSV/XML unreadable | Pre-execution check fails; no worker starts |
| SimplySql missing/unloadable | Build or startup/execution preflight fails; rebuild/correct the image, no startup download |
| Database unavailable | Worker failure and nonzero status logged; scheduler continues |
| Enumeration/import error during scan | Container bootstrap promotes nonterminating PowerShell errors to terminating errors before that snapshot publication |
| Overlap | Wrapper records skip/status 75; existing worker continues |
| Container termination | Worker group is signalled, lock released on exit; review the final log and run metadata before rerunning |

Regular readiness checks require all regular source directories, the actual remote Switch mount, the disc-list CSV/XML, and music root. Requiring music readiness is intentionally conservative because a regular execution may run music based on the existing age decision. Music-only readiness does not require Switch or disc-list sources. Empty source directories are allowed; preflight does not invent row-count thresholds or alter inventory filters.

Mount checks run at startup and before each wrapper execution. They cannot freeze host topology; keep mounts stable during scans. The bootstrap stops observable filesystem/import failures, but a source externally replaced with a valid empty directory or a silent storage fault cannot be inferred from the existing scanner. Do not unmount or replace inventory sources during a scan. This limitation is not addressed by changing Stage 1 inventory behavior.

## Validation commands and remaining host tests

Local tests and build (no production scans):

```bash
python3 -m unittest discover -s tests -v
docker buildx build --platform linux/amd64 --load -t dowehaveit:test .
docker run --rm --entrypoint bash dowehaveit:test -c 'pwsh -NoProfile -Command "Import-Module SimplySql -ErrorAction Stop; Get-Command Open-MySqlConnection"'
docker run --rm --entrypoint python3 dowehaveit:test /opt/dowehaveit/scheduler.py --check-config
```

Deployment checks while frequency is zero:

```bash
docker inspect dowehaveit --format '{{json .Mounts}}'
docker inspect dowehaveit --format '{{json .State.Health}}'
docker logs --tail 100 dowehaveit
docker exec dowehaveit dowehaveit-preflight --mode regular
docker exec dowehaveit dowehaveit-preflight --mode music
```

On an isolated database/source copy, verify manual regular/music runs, published rows/locations/source-count ordering, current-run metadata, transaction behavior, and unchanged retention. Check a failed connection, absent source/remote mount, simultaneous manual/scheduled attempts, logs/retention, and `docker stop` during a scan. Enable frequency 20 and observe actual boundaries and skips during a long scan. Never run a production inventory merely as a build test.

Actual Unraid propagation, recursive read-only flags, source readiness, database connectivity/publication, and Windows UI interaction require the target environments. Windows UI and worker source functions are otherwise unchanged; launch the Windows UI and Time Machine as part of release acceptance.

## Necessary PowerShell changes

Only Common’s Linux database-config selection changes: an existing config beside Common takes precedence over the legacy path. No inventory source path, worker/UI function, SQL, schema, output format, CSV/XML import implementation, or snapshot retention policy changes.

The image’s separate `invoke-worker.ps1` sets `$ErrorActionPreference = 'Stop'` before invoking the existing worker, forwarding `-Music`. This container-only failure policy prevents ordinary nonterminating scan/import errors from silently publishing incomplete rows. Windows and direct noncontainer worker invocations retain their original behavior. Stage 3 physical-disk enumeration is not implemented.
