# CRT.Server deployment

How the contribution API is installed on the AlmaLinux box. Written for the maintainer to run
directly: every step has a **Verify** whose output says plainly whether it worked.

Work through it in order. **Do not skip step 3** - it is the one that makes writing to Production
impossible rather than merely unintended.

> **Nothing in this file is a secret, and nothing in it should become one.** Passwords and
> connection strings are typed on the server, into a file this repository never sees. If a
> credential ever appears in a chat transcript, an email or a commit, treat it as compromised and
> rotate it.

---

## What this deploys

A small ASP.NET Core service that:

- listens on **127.0.0.1 only** and is reached through the existing Apache site at `/api/`;
- runs as a dedicated unprivileged user under `systemd`, restarting automatically;
- in this first deployment answers exactly one endpoint, `GET /api/health`.

The health endpoint is deliberately all there is to begin with. The chain from systemd through
Apache, TLS and SELinux has several independent ways to fail, and diagnosing those at the same time
as diagnosing application code is what turns one evening into one week. Prove the chain with a
service that does nothing, then add features onto a known-good base.

---

## Step 0 - Decisions and prerequisites

You need:

| Thing | Value |
| --- | --- |
| BETA data tree (server-local path) | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data` |
| Production data tree | the matching `app-data` path - **the service must never write here** |
| Web server | Apache, the existing `classic-repair-toolbox.dk` vhost |
| Mail | local postfix on `localhost:25` |
| Port for the service | `5199` on loopback (change it everywhere below if it is taken) |

**Verify the BETA path is what this file assumes:**

```bash
ls -ld /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data
ls -l  /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/dataChecksums.json
```

Both must exist. Note the two are **different directories**: the content lives in `Data/`, while
the manifest the application syncs against sits beside it, one level up. Later phases write the
first and regenerate the second, so both paths matter.

**Verify the port is free:**

```bash
ss -ltn | grep 5199        # no output means free
```

---

## Step 1 - Install the .NET runtime

The service is published framework-dependent, so the box needs the ASP.NET Core runtime (not just
the base .NET runtime, and not the SDK).

```bash
sudo dnf install aspnetcore-runtime-10.0
```

AlmaLinux tracks RHEL, so Microsoft's RHEL feed is the correct one.

**Verify:**

```bash
dotnet --list-runtimes | grep -i aspnetcore
```

Expect a line containing `Microsoft.AspNetCore.App 10.0.x`.

> This is now a component you must keep patched, exactly like Apache and PHP.

---

## Step 2 - Create the service user

```bash
sudo useradd --system --no-create-home --shell /sbin/nologin crt-server
```

**Verify:**

```bash
id crt-server
```

It must exist and belong to no unexpected groups.

---

## Step 3 - File ownership, and the Production interlock

**This is the most important step in the file.**

The service must be able to write the BETA tree (from Phase 4 onward) and must be **unable** to
write the Production tree. Not "configured not to" - *unable*, refused by the kernel. Because you
promote BETA to Production by hand, the service never has any legitimate reason to write
Production, so denying it costs nothing and removes a whole class of accident.

```bash
# A group that owns the BETA content, with the service user in it.
sudo groupadd -f crt-data
sudo usermod -a -G crt-data crt-server

# BETA: group-owned and group-writable, setgid so new files inherit the group
# (that is what keeps Apache serving them normally).
sudo chgrp -R crt-data /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA
sudo chmod -R g+rwX    /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA
sudo find /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA -type d -exec chmod g+s {} \;
```

Production is left exactly as it is - **do not** add `crt-server` to anything that owns it.

**Verify - run all three and read the output carefully:**

```bash
# 1. BETA must be writable by the service user:
sudo -u crt-server touch /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data/.probe \
  && echo "BETA WRITABLE (correct)" \
  && sudo rm -f /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data/.probe

# 2. Production must NOT be writable by the service user:
sudo -u crt-server touch /mydir/http/classic-repair-toolbox.dk/public_html/app-data/.probe \
  && echo "*** PRODUCTION IS WRITABLE - STOP AND FIX THIS ***" \
  || echo "Production refused (correct)"

# 3. Every parent directory must be traversable, or writes fail confusingly later:
namei -l /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data
```

Check 2 **must** print `Production refused (correct)`. If it prints the warning instead, stop here:
the interlock is not in place, and the service would be one configuration mistake away from writing
data that every user downloads.

---

## Step 4 - Publish and copy the binaries

### Where everything lives

Everything this service owns sits under **one root**, beside the site it belongs to rather than
scattered across `/opt` and `/etc`:

```
/mydir/http/classic-repair-toolbox.dk/
├── public_html/                 <- served by Apache
│   ├── app-data-BETA/           <- the service WRITES this (from Phase 4)
│   └── app-data/                <- Production; the service can never write it
└── crt-server/                  <- NOT served by Apache
    ├── app/                     <- binaries + appsettings.Production.json
    └── previous/                <- the last known-good publish, for rollback
```

**`crt-server/` is a SIBLING of `public_html`, never inside it.** Inside, Apache would happily
serve `appsettings.Production.json` - database password and all - to anyone who guessed the URL.
Being outside the document root is what makes that impossible, rather than relying on a `.htaccess`
someone might later reorganise away. Step 9 verifies it directly before any password is typed in.

On the development machine, either run:

```bash
dotnet publish src/CRT.Server/CRT.Server.csproj -c Release -r linux-x64 --self-contained false -o ./publish-server
```

or use Visual Studio's **Publish -> Folder** with Configuration `Release`, Target Framework
`net10.0`, Target Runtime `linux-x64` and Deployment Mode **Framework-dependent** (the last two are
NOT the defaults; getting Deployment Mode wrong produces a ~70 MB self-contained tree that ignores
the runtime installed in step 1). Visual Studio writes to
`bin\Release\net10.0\linux-x64\publish\` rather than `./publish-server`.

**Before copying, check what you have**: about a dozen files, including `CRT.Server.dll`,
`CRT.Data.dll`, `MySqlConnector.dll` and `CRT.Server.runtimeconfig.json`, with no `.exe` and no
sprawl of `System.*.dll`. Hundreds of files means Deployment Mode was left on Self-contained.

**There is also a `Migrations/` SUBFOLDER, and it must travel with the files.** The service reads
its schema scripts from there at startup and refuses to start without it. A copy that flattens the
tree, or a drag-and-drop over a network share that only picks up the loose files, leaves the
service failing with "The migrations directory does not exist".

Copy the files to the box - to your home directory, say - and then, **running the `cp` from the
directory you copied them into**:

```bash
mkdir -p /mydir/http/classic-repair-toolbox.dk/crt-server/app
cp -r ~/publish-server/* /mydir/http/classic-repair-toolbox.dk/crt-server/app/   # adjust to wherever you put them
chown -R root:crt-data /mydir/http/classic-repair-toolbox.dk/crt-server
chmod -R 0750 /mydir/http/classic-repair-toolbox.dk/crt-server
```

The binaries are owned by **root**, so the service cannot rewrite its own program files - an
attacker who compromises the service then cannot make themselves persistent by editing them.

**The group is `crt-data`, matching the `Group=` in the systemd unit in step 5, and that is
load-bearing.** The service process runs with group `crt-data`, so if this directory is group-owned
by anything else the process falls through to the "other" permission bits, which `0750` leaves
empty - it cannot even enter the directory, and systemd fails the unit with the unhelpful
`status=200/CHDIR`. Group `r-x` here lets it read and execute but still not write, which is the
protection intact.

**Verify:**

```bash
APP=/mydir/http/classic-repair-toolbox.dk/crt-server/app

# The deployment is COMPLETE - expect ~12 files, not 1 or 2:
ls -l $APP/ | wc -l
ls -l $APP/CRT.Server.dll $APP/CRT.Data.dll $APP/MySqlConnector.dll $APP/CRT.Server.runtimeconfig.json

# The Migrations SUBFOLDER came across too - the service will not start without it:
ls $APP/Migrations/*.sql

# The service user can enter the directory and read the program, but not write it:
sudo -u crt-server test -x $APP && echo "Directory traversable (correct)" || echo "*** CANNOT ENTER - check group ownership ***"
sudo -u crt-server test -r $APP/CRT.Server.dll && echo "Binary readable (correct)" || echo "*** UNREADABLE - check group ownership ***"
sudo -u crt-server test -w $APP/CRT.Server.dll && echo "*** WRITABLE - fix ownership ***" || echo "Binaries read-only to the service (correct)"
```

All three must print the "(correct)" message. **A copy that silently failed will still pass an
ownership check**, which is why the file count and the three named files are verified first - an
empty or half-populated `app/` otherwise looks fine here and fails confusingly in step 5.

> **`sudo -u crt-server` is WEAKER than the running service, and will pass things the service
> fails.** `sudo -u` gives the user its full supplementary group set (`crt-server` AND `crt-data`),
> while the unit's explicit `Group=crt-data` gives the process `crt-data` alone. So anything owned
> by group `crt-server` reads fine here and is refused at runtime. The checks above are safe
> because everything in `app/` is group `crt-data` - the one group both identities have. Where that
> is not guaranteed, use the `systemd-run` check in step 9 instead.

**Also verify the service directory is not reachable over the web** - this is the one that protects
the password you will put there in step 9:

```bash
curl -s -o /dev/null -w '%{http_code}\n' \
  --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  https://classic-repair-toolbox.dk/crt-server/app/CRT.Server.dll
```

Expect `404`. Anything else - especially `200` - means `crt-server/` is inside the document root
after all, and you must move it before step 9.

---

## Step 5 - The systemd unit

Create `/etc/systemd/system/crt-server.service`:

```ini
[Unit]
Description=Classic Repair Toolbox contribution API
After=network.target

[Service]
# Type=simple, NOT Type=notify. "notify" makes systemd wait for a readiness
# signal over a socket, which a plain ASP.NET Core host never sends - that needs
# the Microsoft.Extensions.Hosting.Systemd package and a UseSystemd() call, and
# this service has neither. With "notify" the process starts and serves requests
# normally while "systemctl start" HANGS until its timeout and then reports the
# unit as failed, which is a thoroughly confusing way to look at a working service.
Type=simple
User=crt-server
Group=crt-data
WorkingDirectory=/mydir/http/classic-repair-toolbox.dk/crt-server/app
ExecStart=/usr/bin/dotnet /mydir/http/classic-repair-toolbox.dk/crt-server/app/CRT.Server.dll
Restart=always
RestartSec=5

Environment=ASPNETCORE_ENVIRONMENT=Production
# Loopback ONLY. The service must never be reachable except through Apache.
Environment=ASPNETCORE_URLS=http://127.0.0.1:5199
Environment=DOTNET_PRINT_TELEMETRY_MESSAGE=false

# Hardening. ProtectSystem=strict makes the whole filesystem read-only to this
# process except the paths named in ReadWritePaths - which is a SECOND, independent
# interlock against writing Production, on top of the file permissions in step 3.
# Do not remove it, and do not add the Production tree to ReadWritePaths.
#
# Note what ReadWritePaths does NOT include: the service's own directory. The
# binaries and appsettings.Production.json stay read-only to the process even
# though they sit under the same site directory as the BETA tree - only the one
# path below is writable, and it is named explicitly rather than inherited.
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=strict
ProtectHome=true
ProtectKernelTunables=true
ProtectControlGroups=true
RestrictAddressFamilies=AF_INET AF_INET6 AF_UNIX
ReadWritePaths=/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now crt-server
```

**Verify:**

```bash
systemctl status crt-server                       # must be active (running)
ss -ltnp | grep 5199                              # must show 127.0.0.1:5199, NOT 0.0.0.0
curl -s http://127.0.0.1:5199/api/health          # {"status":"ok","version":"...","utc":"..."}
```

**Verify it restarts after a crash:**

```bash
sudo systemctl kill -s SIGKILL crt-server
sleep 8
systemctl status crt-server                       # active (running) again
```

**Verify it survives a reboot** (when convenient):

```bash
sudo reboot
# then, once it is back:
curl -s http://127.0.0.1:5199/api/health
```

---

## Step 6 - Apache reverse proxy

### SELinux first - but check whether it is even on

AlmaLinux enforces SELinux by default, and when it does it blocks Apache from making outbound
network connections, including to a local port. Without the boolean below the proxy returns 503 and
the logs read like an application fault.

**Check first**, because this box may not have it enabled at all:

```bash
getenforce
```

- **`Disabled` or `Permissive`** - skip the rest of this sub-step. There is nothing to set, and
  `setsebool` would just fail with "SELinux is disabled". This deployment's box reports `Disabled`.
- **`Enforcing`** - set the boolean:

```bash
setsebool -P httpd_can_network_connect 1
getsebool httpd_can_network_connect        # must print --> on
```

### The vhost

In the vhost for `classic-repair-toolbox.dk`:

```apache
ProxyPreserveHost On
ProxyPass        /api/  http://127.0.0.1:5199/api/
ProxyPassReverse /api/  http://127.0.0.1:5199/api/
RequestHeader set X-Forwarded-Proto "https" "expr=%{HTTPS} == 'on'"
```

**Note the `expr=` on the header, and do not drop it if the vhost serves both :80 and :443.**
This site's vhost is a combined `<VirtualHost *:80 *:443>`, so an unconditional
`RequestHeader set X-Forwarded-Proto "https"` would tell the service that a plain HTTP request
arrived over HTTPS. That is harmless while the only endpoint is health, but from the accounts step
onward the service builds email verification and password-reset links from that header, and a link
built on a lie is a real defect. The `expr=` sets it only when the request genuinely is HTTPS.

**Do not try to solve this by wrapping the block in `<If "%{HTTPS} == 'on'">`.** Apache refuses it
outright - `AH00526: ProxyPass cannot occur within <If> section` - because `ProxyPass` is resolved
at configuration time while `<If>` is evaluated per request. Guarding the header is the way to
express it. (HTTP `/api/` never reaches the proxy on this site anyway: the existing
HTTP-to-HTTPS `RewriteRule` redirects it first.)

**Placement matters.** The site is PHP-driven and may have rewrite rules that catch unknown paths.
The `ProxyPass` must take effect before any such catch-all, or `/api/health` silently returns the
site's own page instead of the service's answer. If that happens, move the block earlier in the
vhost or add an exclusion for `/api/` to the rewrite rules.

**Verify:**

```bash
apachectl configtest && systemctl reload httpd
```

Then test Apache and the network path **separately**, because conflating them wastes time:

```bash
# 1. Apache -> service, bypassing DNS and the router entirely. Same hostname, same
#    SNI, same TLS - but straight to the local interface. This is the real test of
#    whether the proxy configuration works:
curl -i -m 10 --resolve classic-repair-toolbox.dk:443:127.0.0.1 https://classic-repair-toolbox.dk/api/health
```

Expect `HTTP/1.1 200 OK`, `Server: Kestrel` and the JSON body. If you get HTML instead, the PHP
site is swallowing `/api/` - look for an `.htaccess` under the document root, since
`AllowOverride All` is on.

```bash
# 2. The real path a user takes - run this FROM ANOTHER MACHINE, not from the server:
curl -i https://classic-repair-toolbox.dk/api/health
```

**Do not test the public URL from the server itself.** It resolves to the public address and has to
hairpin out to the router and back, which can stall for reasons that have nothing to do with Apache
or the service - a hang there looks exactly like a broken deployment and is not one. Test 1 covers
Apache; another machine covers the network.

**Verify the service is NOT reachable directly from outside** - run this from another machine:

```bash
curl -m 5 http://classic-repair-toolbox.dk:5199/api/health
```

It must fail (refused or timed out). If it answers, the service is bound to a public interface and
`ASPNETCORE_URLS` in the unit is wrong - fix it before going further.

---

## Step 7 - Mail (needed from the accounts step, not for health)

```bash
ss -ltnp | grep :25                        # postfix listening on localhost
echo "test" | sendmail -v your@address
```

Confirm the message actually **arrives**. Deliverability (SPF, DKIM, reverse DNS) is a genuine
prerequisite rather than a detail: if verification mail lands in spam, "email verification is
broken" is indistinguishable from a code fault, and you will debug the wrong thing.

---

## Step 8 - Logs

### Is it working? (the only check you need most of the time)

```bash
systemctl is-active crt-server && curl -s http://127.0.0.1:5199/api/health
```

`active` followed by a line of JSON means yes. Anything else, read on.

### Reading the log WITHOUT the core dump

**A failed start on .NET/Linux writes a ~200-line core dump into the journal**, and the one line
that says what actually went wrong is buried in it. Every command here filters that out. Use these
rather than a bare `journalctl -u crt-server`.

```bash
# What went wrong - ONLY warnings and errors, so a healthy service prints nothing at all:
journalctl -u crt-server -n 20 --no-pager -p warning

# Recent activity, core dump stripped out:
journalctl -u crt-server -n 30 --no-pager | grep -v '^ '

# Follow live (same filter):
journalctl -u crt-server -f | grep -v '^ '
```

The `grep -v '^ '` works because every line of a core dump is a CONTINUATION line beginning with
spaces, while every real log line begins with a timestamp. It is crude and completely reliable.

**`-p warning` is the one to reach for first.** Our own configuration and migration failures are
logged at `crit`, so they survive the filter while systemd's routine chatter does not.

### When you do want everything

```bash
journalctl -u crt-server --since "1 hour ago"
```

> **`journalctl --vacuum-time=...` is NOT per-unit.** Journal files are shared across every unit on
> the box, so `--unit=` filters what you READ, never what gets deleted. Vacuuming to clear one
> noisy service deletes system-wide history for that period. There is no need to vacuum for this
> service at all - the journal rotates itself.

---

## Rolling back

**The binaries.** Before each upgrade, keep the current publish as `previous/`:

```bash
ROOT=/mydir/http/classic-repair-toolbox.dk/crt-server

# Before deploying a new build:
sudo systemctl stop crt-server
sudo rm -rf $ROOT/previous
sudo cp -r $ROOT/app $ROOT/previous
# ... then copy the new publish over $ROOT/app and start again.
```

To roll back:

```bash
sudo systemctl stop crt-server
sudo mv $ROOT/app $ROOT/failed
sudo mv $ROOT/previous $ROOT/app
sudo systemctl start crt-server
curl -s http://127.0.0.1:5199/api/health      # verify before calling it done
```

**`appsettings.Production.json` lives in `app/`, so it travels with a rollback** - which is right,
since a build and the configuration it understands belong together. Copy it across by hand if you
roll back past a change that added a setting.

**Stopping entirely.** Nothing else on the site is affected:

```bash
sudo systemctl disable --now crt-server
# comment out the ProxyPass lines, then:
sudo systemctl reload httpd
```

---

## Troubleshooting

| Symptom | Likely cause | Command that confirms it |
| --- | --- | --- |
| `/api/health` returns the website's HTML | Apache rewrite catches `/api/` before the proxy | `curl -i http://127.0.0.1:5199/api/health` works but the public URL does not |
| Public URL returns 503 | SELinux blocking Apache's outbound connection - only possible if `getenforce` says `Enforcing` | `getenforce`, then `getsebool httpd_can_network_connect` |
| `curl https://<domain>/api/health` hangs when run ON the server | The public domain hairpins out to the router and back; nothing to do with Apache | `curl --resolve <domain>:443:127.0.0.1 https://<domain>/api/health` answers, and so does the same URL from another machine |
| `systemctl start`/`restart` hangs, then reports failed, but the service is actually serving | `Type=notify` in the unit. A plain ASP.NET Core host sends no readiness notification, so systemd waits for one that never comes | `curl -s http://127.0.0.1:5199/api/health` answers while `systemctl` is still hanging. Fix: `Type=simple` |
| `status=200/CHDIR` in `systemctl status` | The service cannot enter `WorkingDirectory`. Its group (`crt-data`, from the unit's `Group=`) does not match the group owning the service directory, so `0750` gives it nothing | `ls -ld <root>/crt-server/app` - the group must be `crt-data`; fix with `chown -R root:crt-data <root>/crt-server` |
| `status=150/EXEC` or "file not found" on start | The publish output was never copied, or only partly | `ls -l <root>/crt-server/app/` - expect ~11 files including `CRT.Data.dll` and `CRT.Server.runtimeconfig.json` |
| `IOException: Permission denied` in `FileConfigurationProvider.Load`, inside `WebApplication.CreateBuilder` | The host cannot open `appsettings.Production.json`. Almost always the group: the file is `root:crt-server` while the process runs `crt-server:crt-data`, because **systemd does not apply supplementary groups when `Group=` is set** | `ls -ln <root>/crt-server/app/appsettings.Production.json` - group must be `crt-data` (not `crt-server`); fix with `chown root:crt-data` on it. Do NOT test with `sudo -u crt-server cat`, which passes regardless - use the `systemd-run` check in step 9 |
| `Result: core-dump`, `signal=ABRT`, and a long stack trace | Normal. An unhandled exception on .NET/Linux exits via `abort()`, so a deliberate configuration refusal looks like a crash | Read the `crit:` line ABOVE the trace - it names the real reason. Use `journalctl -u crt-server -n 20 --no-pager -p warning`, which shows the reason and none of the dump |
| `The migrations directory [...] does not exist` | The `Migrations/` subfolder did not reach the server. A publish from before the folder existed, or a copy that only took the loose files | `ls $APP/Migrations/*.sql`; fix with `cp -r ~/publish-server/Migrations $APP/` then `chown -R root:crt-data $APP/Migrations` |
| `appsettings.Production.json` downloads over HTTPS | `crt-server/` ended up INSIDE the document root | `grep -i DocumentRoot` the vhost; the service directory must be a sibling of `public_html`, not under it. Move it, then rotate the database password - it has been published |
| Service is `activating` then fails | Usually a configuration error - read the message, it names the setting | `journalctl -u crt-server -n 20 --no-pager -p warning` |
| Service runs but Apache 502s | Wrong port in the vhost or the unit | `ss -ltnp \| grep 5199` |
| Writes fail once later phases write data | Group/setgid not applied, or a parent is not traversable | `namei -l <BETA path>` |
| **Approve and publish** answers an error naming a board file and `Permission denied` | The BETA tree's group ownership was lost - **most often by replacing files over a NETWORK SHARE**, which writes them as the share's user and drops both `crt-data` and the setgid bit. The service can create NEW files (the folder is still writable) but cannot OVERWRITE the ones the share replaced, so a publish fails on the first existing image | `ls -l` the named file: the group must be `crt-data` and the mode `-rw-rw-r--`. Re-apply the three commands from step 4 - `chgrp -R`, `chmod -R g+rwX`, and the setgid `find` - **after every bulk copy over the share**, then re-run the publish (re-running is safe and is the documented recovery) |
| Port reachable from outside | `ASPNETCORE_URLS` not loopback | `ss -ltnp \| grep 5199` shows `0.0.0.0` |

---

---

## Step 9 - Configuration (from the second deployment onward)

The first deployment needed no configuration at all. From the build that adds `ServerOptions`, the
service **will not start** without one - by design.

**The configuration file sits in `app/`, beside the binaries.** That is the service's
`WorkingDirectory`, and the ASP.NET Core host reads `appsettings.{Environment}.json` from there
automatically - so this needs no extra systemd setting, no `--contentRoot` and no code change. It
is also why there is no `/etc/crt-server`: a second location would have to be wired up explicitly,
and would leave the service's files in two places for no gain.

```bash
APP=/mydir/http/classic-repair-toolbox.dk/crt-server/app

# Copy src/CRT.Server/appsettings.Example.json across, then:
mv appsettings.Example.json $APP/appsettings.Production.json
chown root:crt-data $APP/appsettings.Production.json
chmod 0640 $APP/appsettings.Production.json
```

**Mode 0640 owned `root:crt-data`**, not 0600 owned by the service user: the service can read its
configuration but cannot rewrite it, which matters if it is ever compromised. The file holds the
database password, so `o+r` must be off.

**The group must be `crt-data`, not `crt-server`, and the difference is not visible on the box.**
The obvious choice is `crt-server` - the file is the service's, the user is `crt-server`, and
`id crt-server` shows that user in both groups, so `root:crt-server` at 0640 looks correct. It is
not: **systemd does NOT apply a user's supplementary groups when the unit sets `Group=`
explicitly**, and step 5's unit sets `Group=crt-data`. The process therefore runs as
`crt-server:crt-data` and nothing else, matches neither owner nor group on a `root:crt-server`
file, falls through to "other", and is refused.

This cost an hour to find, because the obvious check confirms the wrong answer:
`sudo -u crt-server cat` DOES pull in supplementary groups, so the file reads fine by hand while
the service cannot open it. Test with `systemd-run` instead (below), which runs under the unit's
own identity.

**Before typing the password in, prove the file is not web-reachable:**

```bash
curl -s -o /dev/null -w '%{http_code}\n' \
  --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  https://classic-repair-toolbox.dk/crt-server/app/appsettings.Production.json
```

Expect `404`. If it returns `200`, stop: the service directory is inside the document root. Move it
outside before going further.

Then edit it and replace every `/REPLACE/WITH/REAL/PATH/` and `REPLACE_ME`. For this server:

| Setting | Value |
| --- | --- |
| `DataTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data` |
| `ProductionTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data` |
| `ManifestPath` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/dataChecksums.json` |
| `PublicDataBaseUrl` | `https://classic-repair-toolbox.dk/app-data-BETA/Data` |
| `PublicApiBaseUrl` | `https://classic-repair-toolbox.dk/api` |
| `BlobStoreRoot` | `/mydir/http/classic-repair-toolbox.dk/crt-server/blobs` |
| `ConnectionString` | filled in at step 10, when the database exists |

> **`BlobStoreRoot` arrived with Phase 4 and the service will NOT START without it.** A deployment
> of a Phase 4 build over a Phase 3 configuration fails at startup naming this setting - which is
> the no-default rule working, not a fault. Add the line and restart.
>
> It must be **outside `public_html`**, for the same reason `crt-server/` itself is: blobs belong
> to submissions still under review, and anyone who guessed a hash could otherwise download them
> over HTTPS. Putting it beside `app/` keeps everything the service owns under one root.
>
> Create it with the same ownership as the rest of the service directory:
>
> ```bash
> ROOT=/mydir/http/classic-repair-toolbox.dk/crt-server
> mkdir -p $ROOT/blobs
> chown -R crt-server:crt-data $ROOT/blobs
> chmod -R 0750 $ROOT/blobs
> ```
>
> **Note the owner is `crt-server`, not `root`** - unlike the binaries. The service WRITES here,
> so it must own the directory; the binaries are root-owned precisely so the service cannot
> rewrite its own program files.
>
> It also needs adding to the unit's `ReadWritePaths`, or `ProtectSystem=strict` refuses the
> write:
>
> ```bash
> systemctl edit --full crt-server
> # ReadWritePaths=/mydir/http/.../app-data-BETA /mydir/http/.../crt-server/blobs
> systemctl daemon-reload && systemctl restart crt-server
> ```
>
> **This directory grows.** Blobs are kept after a submission is queued, since a reviewer needs
> them; partial uploads are collected automatically after 24 hours. Worth a `du -sh` now and then.

**No systemd change is needed.** The unit already sets `Environment=ASPNETCORE_ENVIRONMENT=Production`
and `WorkingDirectory` to `app/`, which is exactly what makes the host pick this file up. If you
ever move the configuration elsewhere, that is when the unit needs editing - not now.

**Verify - and this is worth doing deliberately, because it demonstrates the interlock rather than
assuming it:**

```bash
# 0. The file is readable AS THE SERVICE ACTUALLY RUNS - same user AND same group as the unit.
#    Do not substitute "sudo -u crt-server cat": that grants supplementary groups the service
#    does not have, so it succeeds on a file the service cannot open.
systemd-run --quiet --wait --pipe --uid=crt-server --gid=crt-data \
  /usr/bin/cat $APP/appsettings.Production.json > /dev/null \
  && echo "Config readable by the service (correct)" \
  || echo "*** SERVICE CANNOT READ CONFIG - check group is crt-data ***"

# 1. With the configuration in place, the service starts:
systemctl restart crt-server
systemctl status crt-server                     # active (running)
curl -s http://127.0.0.1:5199/api/health

# 2. Now REMOVE one required setting (comment out DataTreeRoot), and restart:
systemctl restart crt-server
systemctl status crt-server                     # MUST be failed
journalctl -u crt-server -n 20 --no-pager -p warning   # names the exact setting
```

Step 2 must fail, and the journal must name `CrtServer:DataTreeRoot`. That is the no-default rule
working: a service that was not told where to write data does not start and guess. Put the setting
back and restart before moving on.

---

## Step 10 - The database

The service creates its own schema at startup, so this step only has to create an empty database
and a user for it. Everything else is done by the migration runner.

```sql
CREATE DATABASE crt_review CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
CREATE USER 'crt_review'@'127.0.0.1' IDENTIFIED BY '<generate one on the box>';
GRANT ALL PRIVILEGES ON crt_review.* TO 'crt_review'@'127.0.0.1';
FLUSH PRIVILEGES;
```

**`utf8mb4`, not `utf8`.** MySQL's historical `utf8` is three-byte and silently mangles anything
outside the Basic Multilingual Plane. This database stores contributor display names and free-text
review comments from people all over the world.

**`@'127.0.0.1'`, not `@'localhost'`.** Those are different accounts to MariaDB: `localhost`
matches a unix-socket connection, `127.0.0.1` matches TCP. The service connects over TCP because
that is what the connection string says, so an account created only as `localhost` is refused -
while a `mysql -u crt_review -p` test from the shell succeeds, because that one uses the socket.
That mismatch is why the verify below forces TCP with `-h 127.0.0.1`.

**The grant is database-scoped**, never `ON *.*`. This user has no business reading the other
databases on this box.

**Verify - the `-h 127.0.0.1` is the whole point of this check:**

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "SELECT DATABASE(), CURRENT_USER();"
```

Type the password at the prompt rather than putting it on the command line, where it lands in your
shell history. Expect `crt_review` and `crt_review@127.0.0.1`. If `CURRENT_USER()` says
`@localhost`, the account was created on the wrong host and the service will be refused.

> **If `CREATE USER` fails with `ERROR 1396` and `SELECT user, host FROM mysql.user` shows nothing**,
> the account is still in the server's in-memory privilege cache - which happens when a user was
> removed with a direct `DELETE` rather than `DROP USER`. `FLUSH PRIVILEGES;` clears it. On
> MariaDB 10.4+ the real storage is `mysql.global_priv` and `mysql.user` is only a view over it,
> so check `SELECT User, Host FROM mysql.global_priv WHERE User = 'crt_review';` before concluding
> the row is gone.

Then put the password into `ConnectionString` in `appsettings.Production.json` and restart:

```bash
systemctl restart crt-server
journalctl -u crt-server -n 30 --no-pager | grep -v '^ '
curl -s http://127.0.0.1:5199/api/health
```

The journal must show the migrations being applied:

```
Applying migration 0001 (0001_initial.sql)...
Applied migration 0001.
Applied 1 migration(s).
```

On every later restart it says `Database schema is up to date (1 migration(s) already applied).`
instead. **Verify the tables exist:**

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "SHOW TABLES;"
```

Expect eight: the seven schema tables plus `schema_migrations`.

### Migration 0004 - applied automatically, but worth knowing about

`0004_submission_states_and_system_rows.sql` is the first migration that CHANGES something already
in the database rather than only adding to it: it drops and re-creates
`submissions`' `ck_submissions_state` CHECK constraint.

It exists because the schema and the code disagreed in three ways, none of which any test caught -
every submission test uses an in-memory fake store, so `MySqlSubmissionStore` had never run against
MariaDB at all. Against the real database, **every submission would have failed**:

- `submissions.state`'s CHECK did not allow `uploading` or `abandoned`, which are the first two
  states the code writes;
- `submissions.system_id` is `NOT NULL` with a foreign key to `systems`, and nothing ever inserted
  into `systems`, so the table was empty;
- nothing recorded a system's published revision or content hash.

**Nothing to run by hand** - it applies on the next service start like any other. Verify it landed:

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "SELECT number, file_name, applied_utc FROM schema_migrations ORDER BY number;"
```

Expect four rows, `1` to `4`, ending with `0004_submission_states_and_system_rows.sql`.

> The column is `number`, not `version` - see `MigrationRunner.cs`, which creates this table.

> If the submissions table already holds rows in a state the new constraint rejects, the ALTER
> fails and the service will not start. On a database that has never taken a live submission -
> which is the case here, since submissions would have failed on the foreign key anyway - there is
> nothing to clean up.

### What the migration runner will refuse to do

Worth knowing before it happens, because each of these fails the start rather than warning:

| It refuses when | Because |
| --- | --- |
| An applied migration's file content has changed | An applied migration is history. The database it already ran against cannot be changed retroactively, so editing the file makes two databases disagree while both claim the same version. Revert it and add a NEW migration |
| A number is missing (0001 and 0003, no 0002) | Applying 0003 without 0002 produces a schema no migration sequence can reproduce |
| An unapplied migration is numbered below one already applied | Two branches each added "the next" migration. Applying the loser now runs it out of order |
| A recorded migration's file is gone | An older build was deployed over a newer database |

**Never hand-edit the live schema.** The checksum check exists precisely so that a hand edit
surfaces at the next restart rather than months later.

> **MariaDB commits DDL implicitly**, so a migration that fails halfway leaves behind whatever
> tables it already created. The transaction guarantees only that `schema_migrations` cannot claim
> a migration ran when it did not. If one fails, read the journal, inspect the database by hand,
> and clean up before retrying.

---

## Step 11 - Exercise the accounts endpoints

After redeploying the build that carries them. Run these from the box; each one prints what it got
back, and the expected answer is stated underneath.

```bash
API=http://127.0.0.1:5199/api/accounts
```

**1. Register.** Use an address you can actually read mail at - a verification link is sent to it.

```bash
curl -s -X POST $API/register -H 'Content-Type: application/json' \
  -d '{"email":"you@example.com","password":"a long passphrase here","displayName":"Dennis"}'
```

Expect `202` and a message about mail being on its way.

**2. Register the SAME address again**, with a different password:

```bash
curl -s -o /dev/null -w '%{http_code}\n' -X POST $API/register \
  -H 'Content-Type: application/json' \
  -d '{"email":"you@example.com","password":"a different passphrase","displayName":"Someone Else"}'
```

**Expect `202` again, identical to the first.** This is deliberate and is the anti-enumeration
guarantee: the response must not reveal that the address is taken. Check your mailbox - you should
have a "you already have an account" mail carrying a reset link, addressed to YOUR display name,
not "Someone Else". Confirm no second account was created:

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "SELECT id, email, is_verified FROM accounts;"
```

**3. Verify**, using the link from the first mail (or paste the token):

```bash
curl -s "$API/verify?token=PASTE_THE_TOKEN_HERE"
```

Expect a confirmation message, and `is_verified` now `1` in the query above. Clicking the same
link twice answers politely rather than failing.

**4. Log in:**

```bash
curl -s -X POST $API/login -H 'Content-Type: application/json' \
  -d '{"email":"you@example.com","password":"a long passphrase here"}'
```

Expect `200` and a JSON body with `refreshToken`. Keep it:

```bash
TOKEN=PASTE_THE_REFRESH_TOKEN
```

**5. Who am I:**

```bash
curl -s $API/me -H "Authorization: Bearer $TOKEN"
```

Expect your account, with `isVerified: true` and `isAdministrator: false`.

**6. A wrong password, and an address that does not exist:**

```bash
curl -s -o /dev/null -w '%{http_code}\n' -X POST $API/login \
  -H 'Content-Type: application/json' \
  -d '{"email":"you@example.com","password":"wrong"}'

curl -s -o /dev/null -w '%{http_code}\n' -X POST $API/login \
  -H 'Content-Type: application/json' \
  -d '{"email":"nobody@example.com","password":"wrong"}'
```

**Both must print `401`, with no body explaining which failed.** A different answer for the two
would make this endpoint a way to discover who has an account.

**7. The rate limiter.** Five wrong passwords in a row, then a SIXTH with the CORRECT one:

```bash
for i in 1 2 3 4 5; do
  curl -s -o /dev/null -w "attempt $i: %{http_code}\n" -X POST $API/login \
    -H 'Content-Type: application/json' \
    -d '{"email":"you@example.com","password":"wrong"}'
done

curl -s -o /dev/null -w 'correct password now: %{http_code}\n' -X POST $API/login \
  -H 'Content-Type: application/json' \
  -d '{"email":"you@example.com","password":"a long passphrase here"}'
```

Expect five `401`s and then **`429`** - the correct password is refused too while the limit holds.
That is the point: the limiter guards the password hasher, which allocates ~128 MiB per attempt.
It clears itself after fifteen minutes; there is nothing to unlock by hand.

**8. Refresh, and then reuse the old token:**

```bash
curl -s -X POST $API/refresh -H 'Content-Type: application/json' \
  -d "{\"refreshToken\":\"$TOKEN\"}"
```

Expect `200` and a NEW `refreshToken`. Now present the OLD one again:

```bash
curl -s -o /dev/null -w '%{http_code}\n' -X POST $API/refresh \
  -H 'Content-Type: application/json' -d "{\"refreshToken\":\"$TOKEN\"}"
```

Expect `401` - and note what it did as well as what it answered. Reusing a rotated token means two
parties hold it, so **every session for the account is revoked**, including the new one you just
got. That is deliberate: a re-login is a small price for ending a token theft. Confirm it:

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review \
  -e "SELECT id, revoked_reason FROM sessions; SELECT action, detail FROM audit ORDER BY id DESC LIMIT 5;"
```

You should see every session revoked with `refresh token reused`, and a
`session.reuse_detected` audit row naming the address it came from.

**9. Log in again and log out:**

```bash
curl -s -X POST $API/login -H 'Content-Type: application/json' \
  -d '{"email":"you@example.com","password":"a long passphrase here"}'
# then, with the new token:
curl -s -o /dev/null -w '%{http_code}\n' -X POST $API/logout \
  -H 'Content-Type: application/json' -d '{"refreshToken":"NEW_TOKEN"}'
curl -s -o /dev/null -w '%{http_code}\n' $API/me -H "Authorization: Bearer NEW_TOKEN"
```

Expect `204` then `401` - the session stops working immediately, not at its natural expiry.

### If mail never arrives

The account is still created and the link still exists; only delivery failed. Check postfix rather
than the service:

```bash
journalctl -u postfix -n 30 --no-pager
mailq
```

The service logs a warning and carries on by design - a mail failure must not fail the
registration it accompanies, or the user is left with an account they cannot reach and no way to
retry.

---

## What is deliberately not here yet

**The submission and review endpoints now EXIST but have no step in this runbook**, and that is on
purpose rather than an oversight: each deployment proves one thing at a time, and the steps above
have not yet been run against the real server. Add their steps when you get there, following the
same shape as step 11 - a few `curl` calls with a stated expected answer.

What they will need when you do:

- **`BlobStoreRoot` is a new setting with no default** (Phase 4). The service refuses to start
  without it. It must sit OUTSIDE the document root, and it needs adding to the unit's
  `ReadWritePaths` - see step 9.
- **Migrations 0002, 0003 and 0004** apply automatically on the next start. **0004 is required
  before the server can take a single real submission**; it fixes three schema-versus-code
  disagreements that would otherwise fail every submission against MariaDB.
- **The review endpoints need an administrator account.** `GET /api/review/queue` answers 401
  without credentials and 403 for an account lacking the administrator or reviewer flag, so the
  first useful check is a login followed by a queue request with the bearer token.
- **Publishing writes the BETA tree only.** The service has no write permission on Production (step
  0), which is the interlock working as designed rather than a misconfiguration to fix.

---

## Granting the FIRST administrator

**This is the one step with no code path, and nothing works without it.** `is_administrator`
defaults to 0, no endpoint sets it, and nothing seeds it - so a freshly deployed server has no
account that can approve anything. `GET /api/review/queue` answers 403 for every account, and the
review app shows an empty queue with "this account is not allowed to review submissions".

That is deliberate rather than an omission: an endpoint that grants administrator is an endpoint
that can be abused to grant administrator. The first one is made by hand, on the server, by
somebody who already has database access - which is the maintainer and nobody else.

**Register the account through the normal flow first**, so the password is hashed by the service
rather than written by hand:

```
POST /api/accounts/register   (email, password, displayName)
```

Verify the address as any user would - the verification email goes out through postfix. Then, on
the server:

```sql
-- The email is stored NORMALISED (lowercased, trimmed). Matching on `email` would silently
-- affect zero rows for an address typed with any capital letter.
UPDATE accounts
   SET is_administrator = 1
 WHERE email_normalised = 'you@example.com';

-- Confirm exactly one row, and that it is verified and unlocked - ReviewAuthority refuses an
-- account that is either, whatever its role.
SELECT id, email, is_verified, is_administrator, is_reviewer, is_locked
  FROM accounts
 WHERE email_normalised = 'you@example.com';
```

**No restart is needed.** Authority is resolved per request from the account row, which is the
same property that makes locking an account bite immediately rather than at next login.

**A REVIEWER is granted the same way** with `is_reviewer = 1`. The difference matters: a reviewer
may reject and request changes but **can never approve**, because approving publishes and
publishing is irreversible (no revision history is retained). Give somebody `is_reviewer` when you
want them triaging the queue, and `is_administrator` only when you want them publishing.

## Staying signed in, and how to end a session

The review app remembers its session between launches, so a reviewer signs in once rather than
retyping a long password every time. Two things make that safe, and both are worth knowing when
something looks wrong.

**The session slides rather than being long.** `RefreshTokenDays` (default 30) is measured from
LAST USE, not from login: every authenticated request past the halfway mark pushes `expires_utc`
forward. Somebody using the app weekly is never asked again, while a machine that stops being
used goes cold on its own after one full window. Raising the setting to a year would do the
opposite - it would keep a stolen token alive for a year even if nobody touched it.

**The token is never rotated for this.** Extension is one `UPDATE` to the row already in use, and
the app deliberately never calls `POST /api/accounts/refresh`. That endpoint issues a NEW token
and poisons the old one immediately, so a desktop app that crashed between the two would present
a spent token next launch - and reuse detection would revoke **every session for that account**
and write a `session.reuse_detected` audit entry that reads as an attack. Do not "simplify" the
app into calling refresh.

**On Windows the stored token is encrypted to the logged-in user** (DPAPI). A copy of
`%LOCALAPPDATA%\Classic-Repair-Toolbox\review-session.json` taken to another machine or another
account is inert. **On Linux and macOS nothing is stored at all** - there is no equivalent, and
writing a publishing credential in plain text would be worse than asking for a password - so
reviewers on those platforms sign in each launch, as before.

### Ending a session

| Action | Effect |
| --- | --- |
| **Sign out** in the app | Revokes that one session server-side and deletes the local file. Other machines stay signed in. |
| `UPDATE accounts SET is_locked = 1 ...` | Refused on the **next request**, not at expiry. The app drops to its sign-in screen and forgets the token. |
| `UPDATE sessions SET revoked_utc = NOW() WHERE account_id = ...` | Same, for one account's sessions without locking the account. |

All three surface to the app as a 401, which is what makes it clear the stored token. There is no
state in the app that can outlive the server's answer - the stored expiry is only a hint used to
skip a request that is already known to be dead.

To see who is signed in and when each session was last used:

```sql
SELECT s.id, a.email, s.created_utc, s.last_used_utc, s.expires_utc, s.revoked_reason
  FROM sessions s
  JOIN accounts a ON a.id = s.account_id
 WHERE s.revoked_utc IS NULL
   AND s.expires_utc > UTC_TIMESTAMP()
 ORDER BY s.last_used_utc DESC;
```

`last_used_utc` only became meaningful with sliding expiry - before it, nothing but rotation ever
wrote it, so a session used daily for a month still read as never used.

---

## Trying the review path end to end

Once an administrator exists, in order. Each step's failure tells you something different:

1. **Log in** - `POST /api/accounts/login`. Answers a `refreshToken`; **that value IS the bearer
   token**, despite the name. A client that goes looking for a separate access-token exchange will
   not find one.
2. **Ask for the queue** - `GET /api/review/queue` with `Authorization: Bearer <token>`.
   - `401` - the token is wrong or expired.
   - `403` - the account exists but lacks the role, or is unverified or locked.
   - `200` with an empty list - working, nothing waiting.
3. **Submit something from CRT** against the same server, to put a row in the queue. CRT already
   points at `CrtServerBaseUrl`; no client configuration is needed.
4. **Open it in the review app** - sign in, click the row. The change summary is computed
   server-side, so an empty one here means the payload could not be loaded rather than that
   nothing changed.
5. **Request changes**, then check the contributor side - `GET /api/submissions/{id}` with the
   contributor's upload token returns `reviewerComment`. That sentence is the contributor's ONLY
   feedback; if it comes back empty, the decision did not record.
6. **Approve** something, last. It writes the BETA tree and **cannot be undone**. Check afterwards
   that the board file appeared under its system folder with the `v2.0.0` suffix, that a `.json`
   sidecar sits beside it, and that `systems.current_revision` moved.

**Approving a NEW system needs the tree's master workbooks present.** The generation is read from
them when the system's own folder is empty; with no versioned master anywhere, the publish is
refused rather than writing the frozen unversioned file.
