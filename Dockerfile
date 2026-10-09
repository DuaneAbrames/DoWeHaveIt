# syntax=docker/dockerfile:1
FROM debian:12-slim
ARG POWERSHELL_VERSION=7.6.6-1.deb
ARG SIMPLYSQL_VERSION=2.2.0.106
ENV TZ=America/New_York SCHEDULE_FREQUENCY=20 SCHEDULE_OFFSET=0 LOG_RETENTION_DAYS=30 \
    PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
SHELL ["/bin/bash", "-o", "pipefail", "-c"]
RUN test "$(dpkg --print-architecture)" = amd64 \
    && apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
       mariadb-client curl wget ca-certificates dnsutils iputils-ping iproute2 \
       netcat-openbsd openssl bash coreutils findutils util-linux procps sysstat \
       lsof strace file tree less nano jq tzdata python3 \
    && curl --fail --location --retry 3 https://packages.microsoft.com/config/debian/12/packages-microsoft-prod.deb -o /tmp/microsoft.deb \
    && dpkg -i /tmp/microsoft.deb \
    && apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends "powershell=${POWERSHELL_VERSION}" \
    && rm -rf /var/lib/apt/lists/* /tmp/microsoft.deb
RUN SIMPLYSQL_VERSION="$SIMPLYSQL_VERSION" pwsh -NoLogo -NoProfile -Command \
    '$ErrorActionPreference = "Stop"; Install-PSResource -Name SimplySql -Version $env:SIMPLYSQL_VERSION -Repository PSGallery -Scope AllUsers -TrustRepository; if (-not (Get-Module -ListAvailable SimplySql)) { throw "SimplySql missing" }; Import-Module SimplySql -ErrorAction Stop; Get-Command Open-MySqlConnection -ErrorAction Stop | Out-Null'
COPY container/ /opt/dowehaveit/
COPY tests/ /opt/dowehaveit/tests/
RUN chmod 755 /opt/dowehaveit/entrypoint /opt/dowehaveit/dowehaveit-run /opt/dowehaveit/dowehaveit-preflight /opt/dowehaveit/healthcheck \
    && ln -s /opt/dowehaveit/dowehaveit-run /usr/local/bin/dowehaveit-run \
    && ln -s /opt/dowehaveit/dowehaveit-preflight /usr/local/bin/dowehaveit-preflight \
    && for cmd in pwsh mariadb curl wget dig ping ip nc openssl bash date find flock findmnt mountpoint ionice nice ps top free iostat lsof strace file tree less nano jq python3; do command -v "$cmd" >/dev/null || exit 1; done \
    && test -x /usr/local/bin/dowehaveit-run \
    && python3 /opt/dowehaveit/scheduler.py --check-config \
    && python3 -m unittest discover -s /opt/dowehaveit/tests -v \
    && pwsh -NoLogo -NoProfile -Command '$ErrorActionPreference="Stop"; $t=$null; $e=$null; [void][System.Management.Automation.Language.Parser]::ParseFile("/opt/dowehaveit/invoke-worker.ps1", [ref]$t, [ref]$e); if ($e.Count) { throw $e[0] }; $PSVersionTable.PSVersion.ToString()'
WORKDIR /config
STOPSIGNAL SIGTERM
HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 CMD ["/opt/dowehaveit/healthcheck"]
ENTRYPOINT ["/opt/dowehaveit/entrypoint"]
