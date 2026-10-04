# CRT.Server deployment

How the contribution API is installed on the AlmaLinux box. Written for the project owner to run
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
| Production data tree | the matching `app-data` path - **the service must never write here**, unless you switch on publishing to production (step 13) |
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

> **Since 2026-09-25 this is the DEFAULT, not the only way.** Maintainers can publish a board from
> BETA to Production from CRT's Maintainer tab, once you switch that on - step 13, which
> deliberately undoes part of this step for the Production data folder only. Until you do, this
> step stands exactly as written, and nothing can write Production.

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

### Copying data into BETA (or Production) by hand later

**Whatever you copy in as root arrives owned by root, and the service cannot write into a FOLDER
made that way.** The setgid bit above gives new files and folders the `crt-data` group, but a
normal `cp` as root leaves them without GROUP WRITE (mode 644 files, 755 folders). For FILES that
does not matter since server 3.3.1 - the service replaces a board file by writing a new one beside
it and renaming it into place, which needs only the folder. A FOLDER you copied in, though, cannot
be written into by anything but root.

The service checks every folder before a publish, a production publish or a "Push back to queue"
writes anything, and refuses with the folders named ("nothing was changed") - the log line gives
the exact command. So nothing breaks half-way; you are just asked to do this. **After copying,
run the step 3 commands again for what you copied** (or for the whole tree - it is harmless):

```bash
# The whole BETA folder, as in step 3 - dataChecksums.json sits beside Data/, not in it.
# For Production: .../app-data instead (step 13).
T=/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA
sudo chgrp -R crt-data $T
sudo chmod -R g+rwX    $T
sudo find $T -type d -exec chmod g+s {} +
```

**Do NOT run the service as root to avoid this.** It is the one process on the box that takes
uploads from anybody on the internet; as root, any bug in it - or in a library that reads an
uploaded file - owns the whole server, and step 3's kernel refusal to write Production is gone.

Before 3.3.1 this went wrong in the middle of an approval: the board's workbook was written, the
highlight file beside it (root-owned, copied from Production) was refused, and the maintainer saw
"The server answered 500." with the board half-published. Approving again, after the commands
above, finishes it - the publish writes every file again.

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

**On every LATER deployment, stop the service before the `cp` and start it after:**

```bash
sudo systemctl stop crt-server
cp -r ~/publish-server/* /mydir/http/classic-repair-toolbox.dk/crt-server/app/
chown -R root:crt-data /mydir/http/classic-repair-toolbox.dk/crt-server
chmod -R 0750 /mydir/http/classic-repair-toolbox.dk/crt-server
sudo systemctl start crt-server
```

Copying over a RUNNING service replaces DLLs it has not finished loading; the next time it needs
code from one it reads a file that no longer matches what it loaded, and it can crash with a core
dump whose stack trace carries garbled method names.

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
# A refused setting exits with 78 (EX_CONFIG) and is NOT retried: a restart
# cannot fix a wrong setting, and each retry used to add a core dump to the
# journal every five seconds. From server 3.2.1 a failed database migration
# exits with 78 too, for the same reason. Anything else - the database not up
# yet at boot (exit 69), a crash - still restarts.
RestartPreventExitStatus=78

Environment=ASPNETCORE_ENVIRONMENT=Production
# Loopback ONLY. The service must never be reachable except through Apache.
Environment=ASPNETCORE_URLS=http://127.0.0.1:5199
Environment=DOTNET_PRINT_TELEMETRY_MESSAGE=false

# Hardening. ProtectSystem=strict makes the whole filesystem read-only to this
# process except the paths named in ReadWritePaths - which is a SECOND, independent
# interlock against writing Production, on top of the file permissions in step 3.
# Do not remove it. Add the Production data tree to ReadWritePaths ONLY when you
# switch on publishing to production (step 13), and then only that folder.
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

**The `version` in that JSON is how you tell WHICH build is running** - the thing to check after a
deploy, to confirm the new binaries actually replaced the old ones rather than a copy silently
failing or systemd still holding the previous process. Compare it against the top row of
[VERSION.md](VERSION.md); if it has not changed, the deployment did not land. That version is
maintained by Claude on every server change (see VERSION.md for the SemVer policy), so it moves
whenever the service's behaviour does - you do not have to bump it yourself.

### Reading the log WITHOUT the core dump

**A crash on .NET/Linux writes a ~200-line core dump into the journal**, and the one line that
says what actually went wrong is buried in it. Every command here filters that out. Use these
rather than a bare `journalctl -u crt-server`.

**A refused setting or a failed migration no longer does that** (the migration from server 3.2.1).
The service logs each reason at `crit` and exits with code 78, and `RestartPreventExitStatus=78` in
the unit (step 5) stops systemd retrying, so `systemctl status crt-server` shows `status=78` and the
journal ends with the reasons. A database that does not answer at start exits with 69 instead and
IS retried every five seconds - no core dump, one line each time. A unit
written before 2026-09-25 lacks that line: add it with `sudo systemctl edit --full crt-server`,
then `sudo systemctl daemon-reload`.

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

**That needs server 3.2.1 or later.** From 3.2.1 the service writes its log in systemd's own format
when it runs under systemd - one line per entry, e.g. `CRT.Server.Migrations[0] Applying migration
0013 ...`, with its level recorded in the journal. Before it, each entry was two lines
(`info: CRT.Server.Migrations[0]` and the message below it) and the journal filed EVERY line as
info, so `-p warning` showed only systemd's own lines and never the service's reasons.

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
| `Result: core-dump`, `signal=ABRT`, and a long stack trace | An unhandled exception on .NET/Linux exits via `abort()`: the service crashed. (A refused SETTING and, from 3.2.1, a failed MIGRATION no longer look like this - they exit with `status=78`, see the next row; an unreachable database exits with `status=69` and is retried.) Before 3.2.1 a failed migration and an unreachable database crashed like this too | `journalctl -u crt-server -n 30 --no-pager \| grep -v '^ '` - the reason is the last line before the dump |
| A publish, production publish or push-back is refused: "The server is not allowed to write into [...] in the BETA data, so nothing was changed" | A folder was copied into the tree by hand as root and has no group write (see step 3, "Copying data into BETA by hand later") | `journalctl -u crt-server -n 20 --no-pager -p warning` - the line names the folders and ends with the `chgrp`/`chmod`/`find` command that fixes them; run it, then try again |
| `status=78` and the unit stays stopped | A setting in `appsettings.Production.json` was refused, or (from 3.2.1) a database migration failed. The service logs every reason at `crit` and does not retry, since a restart cannot fix either (`RestartPreventExitStatus=78`, step 5) | `journalctl -u crt-server -n 20 --no-pager -p warning` lists each `Configuration error:` naming its setting, or the migration and the database's own error. Fix it, then `sudo systemctl restart crt-server`. A failed migration may have left tables it created - look before dropping anything |
| `The migrations directory [...] does not exist` | The `Migrations/` subfolder did not reach the server. A publish from before the folder existed, or a copy that only took the loose files | `ls $APP/Migrations/*.sql`; fix with `cp -r ~/publish-server/Migrations $APP/` then `chown -R root:crt-data $APP/Migrations` |
| `appsettings.Production.json` downloads over HTTPS | `crt-server/` ended up INSIDE the document root | `grep -i DocumentRoot` the vhost; the service directory must be a sibling of `public_html`, not under it. Move it, then rotate the database password - it has been published |
| Service is `activating` then fails | Usually a configuration error - read the message, it names the setting | `journalctl -u crt-server -n 20 --no-pager -p warning` |
| Service runs but Apache 502s | Wrong port in the vhost or the unit | `ss -ltnp \| grep 5199` |
| Writes fail once later phases write data | Group/setgid not applied, or a parent is not traversable | `namei -l <BETA path>` |
| **Approve and publish** answers an error naming a board file and `Permission denied` | The BETA tree's group ownership was lost - **most often by replacing files over a NETWORK SHARE**, which writes them as the share's user and drops both `crt-data` and the setgid bit. Since 2026-09-25 a publish writes each file beside its target and RENAMES it into place, which needs write permission on the FOLDER rather than on the old file - so a file that lost its group no longer stops a publish, but a FOLDER that lost it still does | `ls -ld` the named file's folder: the group must be `crt-data` and the mode `drwxrwsr-x`. Re-apply the three commands from step 4 - `chgrp -R`, `chmod -R g+rwX`, and the setgid `find` - **after every bulk copy over the share**, then re-run the publish (re-running is safe and is the documented recovery) |
| **Approve and publish** refuses with "the published tree contains a symbolic link" | A symbolic link sits somewhere between the data root and a file the publish would write. Publishing through it could write outside the data tree, so it is refused and nothing is changed | `find <BETA path> -type l` lists every link. Replace each with the real folder or file |
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
> **This directory grows.** Blobs are kept after a submission is queued, since a maintainer needs
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

## Step 12 - After the security review (2026-09-25)

The review found that a contribution could reach further than intended once approved, and that
nothing bounded what an anonymous sender could store. The code half is in the build; **three
things here need doing by hand on the server**, and the rest is worth knowing.

### 12a - Stop anything in the data trees from ever running (do this first)

The data trees sit inside `public_html`, the site runs PHP, and the vhost has `AllowOverride All`.
The service now refuses dot-files and every file type boards do not use, but a web server that will
execute a `.php` - or obey an `.htaccess` - placed in a data folder is one mistake away from running
contributed code. Close it at the web server, independently of the service.

Add to the `classic-repair-toolbox.dk` vhost, once per tree (BETA and Production):

```apache
<Directory "/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA">
    # No .htaccess in a data tree is ever read, whoever put it there.
    AllowOverride None
    Options -Indexes -ExecCGI -Includes

    # Nothing that could be a script is served at all - refused, not run.
    <FilesMatch "\.(php|phtml|phar|cgi|pl|py|sh)$">
        Require all denied
    </FilesMatch>

    # Browsers must not guess a file's type, and a page served from here may not run script or
    # reach the rest of the site. CRT downloads with its own client and ignores both headers.
    Header set X-Content-Type-Options "nosniff"
    Header set Content-Security-Policy "sandbox"
</Directory>

<Directory "/mydir/http/classic-repair-toolbox.dk/public_html/app-data">
    AllowOverride None
    Options -Indexes -ExecCGI -Includes
    <FilesMatch "\.(php|phtml|phar|cgi|pl|py|sh)$">
        Require all denied
    </FilesMatch>
    Header set X-Content-Type-Options "nosniff"
    Header set Content-Security-Policy "sandbox"
</Directory>
```

`Header` needs `mod_headers` - `httpd -M | grep headers` shows it; AlmaLinux loads it by default.
The site's own PHP pages outside these two folders are not affected.

**Verify - a PHP file in the tree must be refused, not run:**

```bash
BETA=/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA
echo '<?php echo "RAN"; ?>' | sudo tee $BETA/Data/zz-probe.php > /dev/null
curl -s -o /dev/null -w '%{http_code}\n' --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  https://classic-repair-toolbox.dk/app-data-BETA/Data/zz-probe.php      # expect 403
sudo rm -f $BETA/Data/zz-probe.php

curl -sI --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  https://classic-repair-toolbox.dk/app-data-BETA/dataChecksums.json \
  | grep -i -E 'content-security-policy|x-content-type-options'            # expect both
```

Then run a data sync from CRT against BETA and confirm it still completes.

### 12b - Migration 0005 (applies itself; verify it landed)

`0005_binary_system_id_and_upload_budget.sql` makes the system id compare byte for byte in
`systems`, `submissions` and `maintainers`, and adds `submissions.bytes_to_upload`. Before it, the
id ignored case, so an anonymous submission for `commodore/c64/250425` created the row that every
later real submission for that board then attached to - and was rejected.

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "
  SELECT TABLE_NAME, COLLATION_NAME FROM information_schema.COLUMNS
   WHERE TABLE_SCHEMA = 'crt_review' AND COLUMN_NAME = 'system_id';
  SHOW COLUMNS FROM submissions LIKE 'bytes_to_upload';"
```

Expect `utf8mb4_bin` three times and the new column. **It drops and re-creates two foreign keys**,
and MariaDB commits DDL as it goes: if it fails part way, check
`SHOW CREATE TABLE submissions` and `SHOW CREATE TABLE maintainers` for `fk_submissions_system` and
`fk_maintainers_system` before retrying, and re-add whichever is missing with the statements at the
end of the migration file.

**Check for a squatted row** made before the fix - a system id that is a case variant of a real
board folder:

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "SELECT system_id, current_revision, created_utc FROM systems ORDER BY system_id;"
```

Anything whose spelling differs from the published folder names and has no `current_revision` was
never published; reject its submissions in CRT's Maintainer tab.

### 12c - Settings

| Setting | What to do |
| --- | --- |
| `MinimumFreeDiskBytes` | New, defaults to 5 GiB. Submissions and upload chunks are refused (HTTP 507) while the disk holding `BlobStoreRoot` has less than this free, so contributions pause before the site's disk fills. `0` turns it off. |
| `AccessTokenMinutes` | **Removed.** It was validated and read by nothing - the login's `refreshToken` is the one bearer token, with a sliding `RefreshTokenDays` expiry. Delete the line; if it stays, it is ignored. |

### What now happens on its own

* **A submission may only change its own board's folder, `<Manufacturer>/Shared files/` and
  `Generic shared files/`.** Another board's file may be cited only when it is byte-identical to the
  published copy, and is then not written at all. Checked when the submission arrives and again at
  approval.
* **Only file types boards use are accepted** - `.png .jpg .jpeg .gif .bmp .webp .pdf .txt .html
  .htm` - with no dot-files, every file used by a row, and each file's opening bytes matching its
  type. A path that differs from a published one only by capitalisation is refused.
* **Every file is re-verified before a publish writes anything**, and each copy is hashed as it is
  written and renamed into place only on a match. A symbolic link on the way refuses the publish.
* **Per address: at most 20 submissions and 4 GiB of uploads a day.** Maintainers and administrators
  are exempt. Request bodies are capped per route (8 MB for a manifest, an upload chunk or a
  maintainer's saved table, 2 MB for a list of files to remove, 64 KB for everything else).
* **The hourly sweep now also deletes completed blobs no live submission needs**, and clears the
  stored rows of submissions that ended without publishing once they are 30 days old. Blobs of
  merged submissions are kept, so the next edit to that board uploads only what changed.
* **The Maintainer tab lists every file that changes on the server**, not only images, and flags
  shared folders, other boards' files and files no row uses.

**Closing a board to contributions** - the lever for one being flooded - no longer needs a code
change:

```sql
UPDATE systems SET is_accepting = 0 WHERE system_id = 'Commodore/C64/250407';
```

Set it back to `1` to reopen. A board with no `systems` row yet has never been submitted to and is
open.

---

## Step 13 - Publishing to production from CRT's Maintainer tab (optional, 2026-09-25)

**Publishing is two steps now.** Approving a submission publishes it to **BETA**, as before.
Then, once a maintainer has looked at the board in CRT with the BETA data, they open **Beta > Prod**
in CRT's Maintainer tab and publish that board to **Production** - the data every user downloads.
Maintainers can do this for the systems they review; you can do it for all of them. **You are e-mailed
every time a maintainer does it.**

What is copied: every file in the board's BETA folder that Production lacks or has different bytes
for, plus the shared files the board uses that differ. Files the board no longer uses, and that
nothing else in Production uses either, are removed - the plan lists them in red before anyone
approves, and the publish is refused if that list has changed since. If that list holds
a shared file, the board's maintainer AND you must both approve before anything is copied (either of
you first; the other is e-mailed). Each file is hashed as it is copied and only replaces the real
one if it matches BETA. If BETA has changed since the maintainer opened the board, the publish is
refused and they are told to look again.

**It is OFF until you do all of the following.** Until then the Maintainer tab says so, and the
interlock from step 3 holds exactly as before.

**1. Let the service write the Production DATA folder - and only that folder.** This is the one
place this document deliberately undoes step 3. Give the Production `Data` folder the same group
treatment the BETA tree got. The folder above it, with `dataChecksums.json`, needs it too, because
the service rewrites that manifest after every production publish:

```bash
P=/mydir/http/classic-repair-toolbox.dk/public_html/app-data
sudo chgrp -R crt-data $P
sudo chmod -R g+rwX    $P
sudo find $P -type d -exec chmod g+s {} \;

sudo -u crt-server touch $P/Data/.probe && echo "PRODUCTION WRITABLE (now intended)" && sudo rm -f $P/Data/.probe
```

**2. Add it to the unit's `ReadWritePaths`,** or `ProtectSystem=strict` refuses the write:

```bash
systemctl edit --full crt-server
# ReadWritePaths=.../app-data-BETA .../crt-server/blobs .../app-data
systemctl daemon-reload
```

**3. Set all three settings** in `appsettings.Production.json`. All three or none: the service
refuses to start with only some of them, and it refuses any that carries `-BETA` or equals its
BETA twin.

| Setting | Value |
| --- | --- |
| `ProductionDataTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data/Data` |
| `ProductionManifestPath` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data/dataChecksums.json` |
| `ProductionPublicDataBaseUrl` | `https://classic-repair-toolbox.dk/app-data/Data` |

`ProductionTreeRoot` stays as it is; the data folder must sit inside it.

**4. Restart.** Migration 0007 applies itself (three columns on `systems`). The startup check now
also refuses to start if the Production data folder is not writable, which is the point: a missing
permission then shows up as the service not starting, instead of a publish that stops halfway
through a board.

```bash
sudo systemctl restart crt-server
journalctl -u crt-server -n 30 --no-pager
```

**To switch it off again,** empty the three settings and restart, then undo steps 1 and 2 to put
back the kernel-level guarantee.

**The risk you are accepting,** stated plainly: a maintainer's account, with a password as its only
factor (two-factor sign-in is not built), can now publish its own boards to every user. What still
stands in the way: the maintainer only has their own systems, and a shared file also needs your own
approval; only bytes already in BETA can be published; you are e-mailed each time; and the audit trail records who did it.

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
- **The review endpoints need an administrator account, or a maintainer assigned to at least one
  system.** `GET /api/review/queue` answers 401 without credentials and 403 for an account with
  neither, so the first useful check is a login followed by a queue request with the bearer token.
- **Migration 0006** (Phase 6 roles, 2026-09-25) applies itself on the next start: it renames the
  pool table to `reviewers`, drops `accounts.is_reviewer`, and adds
  `submissions.touches_shared_files`. Migration 0010 later renames the table back to `maintainers`
  (see "The reviewer role is now called maintainer" below). Nothing to do by hand.
- **Publishing writes the BETA tree only.** The service has no write permission on Production (step
  0), which is the interlock working as designed rather than a misconfiguration to fix. Publishing
  to production from CRT's Maintainer tab is a separate, switched-off-by-default feature - step
  13.

---

## Granting the FIRST administrator

**This is the one step with no code path, and nothing works without it.** `is_administrator`
defaults to 0, no endpoint sets it, and nothing seeds it - so a freshly deployed server has no
account that can approve anything. `GET /api/review/queue` answers 403 for every account, and the
Maintainer tab shows an empty queue with "this account is not allowed to review submissions".

That is deliberate rather than an omission: an endpoint that grants administrator is an endpoint
that can be abused to grant administrator. The first one is made by hand, on the server, by
somebody who already has database access - which is the project owner and nobody else.

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
SELECT id, email, is_verified, is_administrator, is_locked
  FROM accounts
 WHERE email_normalised = 'you@example.com';
```

**No restart is needed.** Authority is resolved per request from the account row, which is the
same property that makes locking an account bite immediately rather than at next login.

## Granting MAINTAINERS - from CRT's Maintainer tab, not SQL

**A maintainer is somebody you assign to a system, and that person reviews AND publishes changes
to exactly the systems you assign** (the project owner's two-role model, 2026-09-25). There is no
flag to set: sign in on CRT's Maintainer tab as the administrator, open the **Systems** screen,
pick a system, and under its maintainers either:

- **choose somebody who has an account** from the list and press **Add as maintainer** - the list
  says, in the line itself, why one cannot be granted (address not verified, locked, or already an
  administrator); or
- **type the email address of somebody new** and press **Send invitation** (2026-09-27). They get a
  mail with a code, and the mail tells them the rest: install CRT (or use the one they have), tick
  **"Enable Maintainer tab"** in its Configuration tab, choose **"I have an invitation"** on that
  tab's sign-in screen, paste it, and pick a name and a password. That makes their account - already verified - and they
  maintain the system from that moment. The code works for 14 days and once. Until it is used the
  invitation is listed under the system's maintainers with a **Withdraw** button; inviting the same
  address again sends a new code and stops the old one. An address that already has an account is
  not invited - choose it from the list instead.

Each maintainer has a **Remove** button beside them. Migration 0012 (`maintainer_invitations`)
applies itself on the next start; check it with
`mysql -u crt_review -p -h 127.0.0.1 crt_review -e "SHOW TABLES LIKE 'maintainer_invitations';"`.

Every board in the BETA tree is listed, whether or not anything has ever been submitted to it,
so a maintainer can be assigned before the first contribution arrives. **Removal takes effect on
the person's very next request** - authority is read from the `maintainers` table on every call,
never cached in a session.

What a maintainer gets: the queue filtered to their systems, and Approve on each. One approval
publishes - theirs or yours - with ONE exception (since server 3.1.0, 2026-09-27): a submission that
REPLACES a file that already exists under `Shared files` or `Generic shared files` with different
content needs TWO approvals, the maintainer's AND yours, for BETA and again for Production, because
it changes what every board using that file shows (it says "replaces a shared file" in the row).
Adding a NEW shared file needs only the one approval. Either of you may approve a replacement first;
that publishes nothing, the row then says "one of two approvals given", and the other is e-mailed.
The second approval publishes. On a board with no maintainer, your approval alone does it. When a
submission is queued, whoever must approve is e-mailed: its system's maintainers, plus you on a
shared-file replacement; with nobody assigned, you alone.

**A publish, a promotion or a push-back removes files only inside the system's own folder**
(`Commodore/C64/250407/...`) - never under `Shared files`, `Generic shared files` or another
system's folder. A shared file no board uses any more stays where it is; it shows up in **Account >
Unused files** (the Maintainer tab's Admin screen when this was written), where you remove it when
you choose.

Migration 0008 (the two approval tables) applies itself on the next start, like 0006 and 0007.

Only an administrator sees these controls or can call `/api/admin/*`; the server refuses everyone
else regardless of what the app shows. Accepting an invitation (`/api/accounts/accept-invitation`)
needs no sign-in - the code is the proof.

### Changing a submission before publishing it

**View in table format** in the maintainer application (now CRT's Maintainer tab, where the table
simply opens with the submission) shows a submission as the Drafts tab's table,
coloured against the published board, and a maintainer of that board (or you) can correct rows there
and press Save changes. That saves a new version of the submission: the contributor's original is
kept in the database, any approval already given is cleared (it was given to other content), the
contributor is told in their mail and in CRT, and the audit trail records who changed it. A row may
only point at a file the submission carries or one already published - new files still come from
contributors. Migration 0009 (`submission_amendments`) applies itself on the next start.

### The reviewer role is now called maintainer (migration 0010)

The review application is now **CRT Maintainer**, and the role it serves is **maintainer** - it was
"reviewer" until 2026-09-25. (**Since 2026-09-29 there is no separate CRT Maintainer application**:
it became the Maintainer tab in CRT, shown by ticking "Enable Maintainer tab" in CRT's Configuration
tab. The server did not change for that beyond the wording of its mails - 3.5.1. The notes below that
say "deploy with the new CRT Maintainer" describe what was true when each build shipped.) Migration 0010 applies itself on the next start: it renames the pool
table `reviewers` back to `maintainers`, rewrites the stored approval roles (`reviewer` becomes
`maintainer`, with the CHECK constraints that allow them) and the grant/revoke actions in the audit
trail. Nothing to do by hand. Check it landed:

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "
  SHOW TABLES LIKE 'maintainers';
  SELECT DISTINCT role FROM submission_approvals;
  SELECT DISTINCT role FROM production_approvals;"
```

The first answers one row; the other two list only `maintainer` and `administrator` (or nothing).

**Deploy this server build before anyone uses the new CRT Maintainer**: the new app calls
`/api/admin/maintainers`, which the previous server does not have, and the old review application
calls routes this server no longer answers.

**If the service stops at startup during 0010**, fix what the log names and restart: every statement
in 0010 is safe to run again except the final rename, which only runs once everything before it has
succeeded. In the unlikely case the log says `Table 'reviewers' doesn't exist`, the rename happened
but its record did not; put the name back and restart, and 0010 runs again from the top:

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "RENAME TABLE maintainers TO reviewers;"
```

### Unused files

A publish now removes the files a board stops using, when nothing else uses them either - the
maintainer sees that list before approving. Files that were ALREADY unused are yours: **Unused files**
under **Admin** in CRT's Maintainer tab (administrator only) lists them per data tree, BETA or
Production, with sizes. Look through the list, tick the box, and press Remove. The server removes
only the files on the list that it still finds unused, rewrites that tree's `dataChecksums.json`,
and records every file in the audit trail (`data.unused_removed`).

The first time, BETA should show the files the repository copy of the data no longer has (the 50
removed on 2026-09-25) plus anything else that has gathered on the server. If the list says it
cannot be made - a master workbook listing a board file that is not there, or a workbook that cannot
be read - nothing can be removed until that is fixed, and no publish removes anything either.

A user's CRT keeps its downloaded copy of a removed file unless "Delete orphan and non-used files"
is switched on in its Configuration tab (off by default).

### A new system's place in the drop-down lists (migration 0011, server 2.0.0)

A new system now has to be PLACED in CRT's drop-down lists before it can be approved: its hardware
name, board name and hardware notes, and which row of the main Excel data file it goes after. A
maintainer of the system (or you) does that on the **Systems** screen of CRT's Maintainer tab, by dragging
it into place in the full list. An approval of a new system nobody has placed is refused with a
message saying so; the review screen warns about it above the table first.

What happens to the main Excel data file (`Classic-Repair-Toolbox.v<newest>.xlsx`, in each tree):

- **Publishing to BETA** inserts the system's row into BETA's file, after the row it was placed
  after, before the submission is marked merged - so CRT users on the BETA source see it at once.
- **A system already in BETA but not in its list** (copied there by hand, or merged before this
  version) is added to BETA's list the moment it is placed, and the BETA manifest is rebuilt.
- **Publish to production** inserts the same row into production's file, straight after the nearest
  row above it that production also lists. It is refused, with nothing copied, when neither file
  lists the system.
- **Push back to queue** on a system that never reached production takes its row out of BETA's file
  again. Its placement is kept, so publishing it again puts it back in the same place.
- Any of these is refused when another system is already listed under the same hardware name AND
  board name - CRT keys a board by that pair, so it would show one board twice.

Only the one row changes; the other sheets (Oscilloscope) and every other row are left exactly as they
were, formatting included. Older generations of the file (`v1.x`) are never touched.

Migration 0011 applies itself on the next start: six nullable `listing_*` columns on `systems`,
where each placement is kept. Nothing to do by hand. Check it landed:

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "SHOW COLUMNS FROM systems LIKE 'listing_%';"
```

It answers six rows.

### One submission in BETA per system, and a system's history (server 3.0.0)

**A system takes one submission into BETA at a time** (owner decision, 2026-09-27). While a
system has a submission in BETA that has not been published to production, no other submission of
that system can be approved: the Maintainer tab greys Approve out with the reason above the table, and
the server refuses it anyway. Publish the one in BETA to production (**Beta > Prod**) or push it
back to the queue, and the next can go. This is because pushing back always takes back everything
merged since the last promotion - a publish replaces the board's rows, so one contributor's work
cannot be picked out of several. The rule is on only when production publishing is configured
(step 13); without it nothing ever leaves BETA.

Each system on the **Systems** screen now shows its **History**, newest first: submissions sent and
decided (and by whom), promotions to production, push-backs, maintainers added, removed and
invited, and its place in the drop-down lists being saved. It is read from the submissions and the
audit trail, so there is nothing to migrate - events from before this version appear as far as the
audit trail recorded them (a maintainer removed before it shows as "account N" rather than an
address).

**Deploy this server build together with the new CRT Maintainer.** An older CRT Maintainer shows the
refusal's reason on Approve, but it has no Systems-screen placement, so a new system cannot be
approved from it at all.

### A board copied to production by hand, and "taken back out of BETA" (migration 0015, server 3.6.0)

**A board you copy into production by hand is no longer "waiting for production".** The server
only knew a board had reached production when "Publish to production" put it there, so a board
copied as root stayed listed under **Beta > Prod** and - through the one-in-BETA rule above - blocked
every new approval of that system. Now, when the record says a system waits, the server compares
the two trees: if production already holds exactly what BETA has (nothing to copy, nothing to
remove, the system already in production's drop-down lists), it records the system as in production
and says so in the system's **History** ("Found already in production"). Nothing for you to do; the
Beta > Prod list re-checks a system that genuinely differs at most every ten minutes.

**"Taken back out of BETA" is now recorded, not guessed.** Migration **0015** adds the table
`submission_beta_returns`, and applies itself on the next start like every migration. A push-back
writes one row per submission it returns. Submissions pushed back BEFORE this version have no row,
so they now read as plain "Waiting for review" to their contributor; the maintainer's comment is
still shown. Check it landed:

```sql
SHOW TABLES LIKE 'submission_beta_returns';
```

## Board views - which boards CRT users look at (migration 0013, server 3.2.0)

From the CRT release that carries it, CRT counts a **view** every time a published board has been on
screen for 10 seconds, and sends its views here in batches (`POST /api/usage/board-views`, anonymous,
like the launch check-in). The server stores **one row per view** in `crt_review.crt_board_views`:
the board and its names from the published main Excel data file, when, the CRT version, operating
system and CPU, whether CRT was downloading from the BETA source, and the **country** - looked up
from the sender's address when the batch arrives. **No address and no identifier is stored.** The
launch check-in (`crt_update`) carries on beside it - received by this service too since 3.11.0,
which replaced the `app-checkin` PHP page.

Nothing to configure (one optional setting, `CountLocalNetworkBoardViews`, is described below). Four
things to do once, each checkable on its own:

**1. Deploy as usual; migration 0013 applies itself.** Check both tables exist:

```bash
mysql -u root -p -e "SHOW TABLES FROM crt_review LIKE 'crt\_board%';"
```

Expect `crt_board_view_batches` and `crt_board_views`. The first only remembers which batches have
arrived (for 60 days), so a batch CRT sends again is not counted twice.

**2. The country lookup is an outbound call to ip-api.com** - the service the check-in has always
used, over plain HTTP, three seconds at most. Check the service's user can reach it:

```bash
sudo -u crt-server curl -s -m 5 "http://ip-api.com/json/8.8.8.8?fields=status,countryCode,country"
```

Expect `{"status":"success","country":"United States","countryCode":"US"}`. If it fails, views are
still stored, only without a country, and the journal says `Looking up a sender's country failed`
(without the address). The unit's `RestrictAddressFamilies` already allows it.

**3. Let the Fun facts page read the new table** (the page reads the database directly, as its other
charts do; the `helligsoe` user has table-level grants only - never `crt_review.*`):

```sql
GRANT SELECT ON crt_review.crt_board_views TO 'helligsoe'@'192.168.20.%';
```

**4. Copy the two Fun facts files** from `Assets/Webserver/funfacts/` to
`classic-repair-toolbox.dk/public_html/funfacts/`: the new `funfacts_board_views.php` (two charts:
the most viewed boards and views per country, both the last 30 days, BETA-source views left out) and
`funfacts.php`, which now ends its charts with `require "funfacts_board_views.php";`. Before step 3 or
before the first views arrive, the two charts are simply empty.

**Until board views go live, only your own PC (`192.168.30.11`) sees the two charts** - everybody
else gets the Fun facts page exactly as before. The check is the four lines marked `NOT LIVE YET` at
the top of `funfacts_board_views.php`; delete them (locally and on the server) to show the charts to
everyone.

**Verify** once somebody on the new CRT has had a board open for 10 seconds and a minute has passed:

```bash
mysql -u root -p -e "SELECT viewedUtc, hardwareName, boardName, version, countryCode, fromBeta, fromLocalNetwork FROM crt_review.crt_board_views ORDER BY id DESC LIMIT 5;"
```

**Your own CRT at home is counted too - for now** (your request, 2026-09-27, so the numbers can be
checked). A batch from your own network arrives with a `192.168.*` address, which places nobody, so
its views get the country of the server's own public address (the same ip-api.com call, asking about
the server itself) and are **marked `fromLocalNetwork = 1`**. Each home batch is also said in the
journal:

```bash
journalctl -u crt-server --since "10 min ago" --no-pager | grep "board views"
```

Expect `A batch of board views from a local-network address arrived: N stored, 0 not stored.` To
list only your own views:

```bash
mysql -u root -p -e "SELECT viewedUtc, hardwareName, boardName, countryCode FROM crt_review.crt_board_views WHERE fromLocalNetwork = 1 ORDER BY id DESC LIMIT 10;"
```

**When you want home views to stop counting**, add this line inside the `CrtServer` section of
`appsettings.Production.json` and restart the service (`sudo systemctl restart crt-server`):

```json
"CountLocalNetworkBoardViews": false,
```

From then on a home batch is accepted and nothing of it is stored - the check-in's rule - and the
journal says `0 stored, N not stored`. The views already stored stay. To remove them as well:

```sql
DELETE FROM crt_review.crt_board_views WHERE fromLocalNetwork = 1;
```

**Rate limit:** one address may send 60 batches an hour (in memory only - the address is not
written anywhere); more are answered `429`, and CRT keeps those views and sends them later. **CRT
Maintainer's Systems screen** shows each system's views ("48 views in 30 days" on its line, and the
counts, countries and BETA-source views in its panel), so deploy the new CRT Maintainer with this
build to see them - an older one simply shows none.

## API usage - which CRT versions call which route (migration 0018, server 4.7.0)

From server 4.7.0 every request that reaches a route is counted in memory - per day, route PATTERN
(`/api/review/submissions/{submissionId:long}`, never a real path) and the CRT version its User-Agent
names - and written into `crt_review.crt_api_calls` every five minutes and when the service stops.
**No address, no account and no User-Agent text** is stored. CRT's Maintainer tab shows it under
Account > **API usage**: every route the server maps, the versions that called it
in the last 30, 90 or 365 days, and how many installations launched each version (counted from
`crt_update`, inside the database).

Nothing to configure. **Deploy as usual; migration 0018 applies itself.** Check the table exists:

```bash
mysql -u root -p -e "SHOW TABLES FROM crt_review LIKE 'crt\_api\_calls';"
```

and, five minutes after the first CRT has talked to the server:

```bash
mysql -u root -p -e "SELECT callDate, method, route, version, calls FROM crt_review.crt_api_calls ORDER BY lastUtc DESC LIMIT 10;"
```

What the numbers are for - retiring a route only old versions still call - is CLAUDE.md's
"Installed CRTs keep working". A route is never simply removed: it stays and answers those versions
"please update CRT".

## Going live - resetting the contribution data (server 4.7.0)

Account > **Reset contribution data** (in CRT's Maintainer tab) deletes everything people sent and did
through the contribution service during testing: every submission (and, on disk, its uploaded
files), every account that is NOT an administrator - so every maintainer and invitation goes with
them - every system record, the whole history, the board views and the API usage counts. **It never
touches the BETA or stable data**, their main Excel data files or their checksum manifests, the
launch check-ins (`crt_update`) or the saved feedback. Your administrator account, and your session,
stay.

**It works only while the server allows it.** The setting is off by default, so a click alone - or
a stolen administrator session - can never wipe the database. The order for going live:

1. **Make BETA the same as stable.** Copy the stable data over the BETA data by hand, as in "Copying
   data into BETA (or Production) by hand later" under step 3 - including removing from BETA what
   stable does not have, if you want BETA exactly level.
2. **Rebuild the manifests**: Account > Rebuild checksum manifests.
3. **Switch the reset on.** In the `CrtServer` section of `appsettings.Production.json`:

   ```json
   "AllowDataReset": true,
   ```

   then `sudo systemctl restart crt-server`.
4. **Reset.** Account > Reset contribution data shows what will go; type `RESET`
   and press the button. It answers with what was deleted and reads the counts again - all zero
   except one history entry (the reset itself) and your administrator account.
5. **Switch it off again**: set `"AllowDataReset": false` (or remove the line) and restart.

Check the journal says so:

```bash
journalctl -u crt-server --since "10 min ago" --no-pager -p warning | grep "reset the contribution data"
```

**Testers' CRTs** still remember what they sent. Their submissions now answer "not found", which CRT
shows as **"No longer on the server"** - and the same draft can be submitted again. Nobody is mailed.

## Feedback from CRT - replacing the app-feedback PHP page (server 3.10.0)

From server 3.10.0 the service receives what CRT's Feedback tab sends (`POST /api/feedback`) and
does what `public_html/app-feedback/index.php` did: the attached files are saved in a
`feedback-<random>` folder under `/mydir/http/classic-repair-toolbox.dk/user-feedback` - the same
folder, so your network share keeps working - and you get a mail with the text, the log file and the
settings in it, its "Internal reference" naming the folder. **What is different:**

- the mail is **HTML**, with the log, settings and traces in a fixed-width font (and a plain-text
  copy beside it, as every mail now has);
- it comes **from `MailFromAddress`, with Reply-To the user's address**. The PHP sent it from the
  user's own address, which fails the user's provider's SPF check and tends to land in spam. Reply
  as usual and it goes to the user;
- the crash log is shown in the mail too, like the log, rather than saved as a file;
- a log too large to show (over 3 MB) is saved with the files instead, and the mail says so;
- the attachments may be **at most 250 MB packed** - CRT says so before it sends anything - and one
  address may send **10 feedbacks an hour**. A zip may hold at most 40,000 entries (server 4.7.2 -
  more is refused before it is opened, and the mail says so) and unpack to at most 2 GB and 20,000
  files, and an upload is refused while the disk would drop below `MinimumFreeDiskBytes` (5 GB);
- the saved feedback folders may take **at most `FeedbackMaxStoredBytes` (20 GB) all together**
  (server 4.3.1). Past it, the text is still mailed but the files are not saved, and the mail says
  so - delete old `feedback-*` folders to make room; the service counts the folder again every
  five minutes (server 4.7.2), so room made through the share counts within that. Optional; `0`
  turns the total off;
- the saved folders and files are **group-writable**, so you can delete them through your share
  (steps 1 and 5);
- a zip entry naming a path outside its folder (`../`) is never written, and the mail says it was
  skipped.

**The service will not start until steps 1 to 3 are done** - two new settings are required. Do them
before deploying 3.10.0, or straight after, before the restart.

**1. Hand the feedback folder to the service** - owner and group `crt-server` (existing `feedback-*`
folders are left as they are). Step 2's `useradd --system` made a `crt-server` GROUP beside the
user; the first line checks, and makes it if it is missing. It is the group, not `crt-data`, because
step 5 may put your share's user in it - and that user must not get write access to the BETA data
along with it. Nothing else is in the `crt-server` group.

```bash
getent group crt-server || sudo groupadd crt-server

F=/mydir/http/classic-repair-toolbox.dk/user-feedback
sudo chown crt-server:crt-server $F
sudo chmod 2775 $F          # rwxrwsr-x: the setgid "s" makes everything inside get the group

# Must print FEEDBACK WRITABLE (correct):
sudo -u crt-server touch $F/.probe && echo "FEEDBACK WRITABLE (correct)" && sudo rm -f $F/.probe
```

Every folder and file the service saves is made group-writable (`rwxrwsr-x` / `rw-rw-r--`) and gets
the folder's group, `crt-server`, so anybody in that group can delete it - see step 5.

From here until step 6, the old PHP page can still mail feedback but can no longer save attached
files into the folder (it runs as `apache`) - its mail then says so. Do steps 2 to 6 in one go.

**2. Add the folder to the unit's `ReadWritePaths`,** or `ProtectSystem=strict` refuses the write:

```bash
systemctl edit --full crt-server
# ReadWritePaths=.../app-data-BETA .../crt-server/blobs /mydir/http/classic-repair-toolbox.dk/user-feedback
# (keep .../app-data on the line too if you switched on publishing to production, step 13)
systemctl daemon-reload
```

**3. Add the two settings** inside the `CrtServer` section of `appsettings.Production.json`:

```json
"FeedbackRoot": "/mydir/http/classic-repair-toolbox.dk/user-feedback",
"FeedbackToAddress": "dennis@classic-repair-toolbox.dk",
```

Then deploy 3.10.0 as usual and restart (`sudo systemctl restart crt-server`). If it refuses to
start, the journal names the setting (`journalctl -u crt-server -n 20 --no-pager -p warning`).

**4. Try it from the server, through Apache:**

```bash
# Text only - expect HTTP 200, "Success" and Server: Kestrel, and a mail a moment later:
curl -i -m 30 --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  -F "feedback=Test from the server" -F "email=dennis@classic-repair-toolbox.dk" -F "version=deploy-test" \
  https://classic-repair-toolbox.dk/api/feedback

# With a file - expect a new feedback-* folder holding note.txt, named in the mail:
cd /tmp && echo "hello" > note.txt && python3 -m zipfile -c feedback-test.zip note.txt
curl -i -m 30 --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  -F "feedback=Test with a file" -F "version=deploy-test" \
  -F "attachmentFile=@/tmp/feedback-test.zip;type=application/zip" \
  https://classic-repair-toolbox.dk/api/feedback
ls -l /mydir/http/classic-repair-toolbox.dk/user-feedback | tail -3
```

The journal says each one, without the address or the text:
`Feedback received and mailed: feedback-..., 1 file(s) saved.`

**5. Check you can delete feedback through your share.** Delete the test folder from step 4 in
Windows, through the share, as you would old feedback.

- **It is deleted** - nothing more to do (your share works as root, or as a user already in the group).
- **Access denied** - put the share's user in the `crt-server` group. To find that user, list the shares:

  ```bash
  testparm -s 2>/dev/null | grep -iE '^\[|path *=|force user'
  ```

  Find the share whose `path` holds `user-feedback`. If it has a `force user = X` line, the user is
  `X`; if not, it is the name you log in to the share with. Then:

  ```bash
  sudo usermod -a -G crt-server X       # X = that user
  sudo systemctl restart smb            # a share only picks up a new group on a new connection
  ```

  Reconnect the share in Windows and delete the folder again.

**6. Send the old address to the service, for CRTs already installed.** Every CRT up to this release
posts to `https://classic-repair-toolbox.dk/app-feedback/`. In `/etc/httpd/conf/httpd.conf`, at the
bottom of the `classic-repair-toolbox.dk` `<VirtualHost>`, are the `/api/` lines from Step 6 - add
the two `/app-feedback/` lines right after `ProxyPassReverse /api/`, so the end of that vhost reads:

```apache
  ProxyPreserveHost On
  ProxyPass        /api/  http://127.0.0.1:5199/api/
  ProxyPassReverse /api/  http://127.0.0.1:5199/api/
  ProxyPass        /app-feedback/  http://127.0.0.1:5199/api/feedback
  ProxyPassReverse /app-feedback/  http://127.0.0.1:5199/api/feedback
  RequestHeader set X-Forwarded-Proto "https" "expr=%{HTTPS} == 'on'"
</VirtualHost>
```

```bash
apachectl configtest && systemctl reload httpd

# The old address now answers from the service - expect "Success" and Server: Kestrel:
curl -i -m 30 --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  -F "feedback=Test through the old address" -F "version=deploy-test" \
  https://classic-repair-toolbox.dk/app-feedback/
```

If that answer comes from the PHP page instead (no `Server: Kestrel`), the `ProxyPass` lines are
after something that catches the path - move them up beside the `/api/` lines.

**Apache must let a 256 MB request through.** Its own default allows 1 GB, but a `LimitRequestBody`
set lower anywhere in the configuration would refuse large feedback before the service sees it:

```bash
grep -rin "LimitRequestBody" /etc/httpd/ 2>/dev/null
```

Nothing printed, or only values of `268435456` or more (or `0`, unlimited), is fine. A smaller
value: raise it to `268435456` where it is set, then `apachectl configtest && systemctl reload httpd`.

**7. Retire the PHP page.** Once step 6 answers from the service, move the page out of the web
root (kept rather than deleted, in case you want it back):

```bash
sudo mv /mydir/http/classic-repair-toolbox.dk/public_html/app-feedback /root/app-feedback.retired-2026-10-03
```

Leave `public_html/libraries/` alone - `app-contribution` uses it (PHPMailer) until it is retired too.

**Going back**, should it ever be needed: move the folder back, remove the two `ProxyPass` lines and
reload Apache. The service's route can stay; nothing else uses it.

## Staying signed in, and how to end a session

CRT's Maintainer tab remembers its session between launches, so a maintainer signs in once rather
than retyping a long password every time. (It is the same file, in the same folder, the separate CRT
Maintainer application kept until 2026-09-29, so a maintainer who was signed in there is still signed
in on the tab.) Two things make that safe, and both are worth knowing when
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
`%LOCALAPPDATA%\Classic-Repair-Toolbox\maintainer-session.json` taken to another machine or another
account is inert. **On Linux and macOS nothing is stored at all** - there is no equivalent, and
writing a publishing credential in plain text would be worse than asking for a password - so
maintainers on those platforms sign in each launch, as before.

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
4. **Open it in CRT's Maintainer tab** - sign in, click the row. The change summary is computed
   server-side, so an empty one here means the payload could not be loaded rather than that
   nothing changed.
5. **Request changes**, then check the contributor side - `GET /api/submissions/{id}` with the
   contributor's upload token returns `maintainerComment`. That sentence is the contributor's ONLY
   feedback; if it comes back empty, the decision did not record.
6. **Approve** something, last. It writes the BETA tree and **cannot be undone**. Check afterwards
   that the board file appeared under its system folder with the `v2.0.0` suffix, that a `.json`
   sidecar sits beside it, and that `systems.current_revision` moved.

**Approving a NEW system needs the tree's master workbooks present.** The generation is read from
them when the system's own folder is empty; with no versioned master anywhere, the publish is
refused rather than writing the frozen unversioned file.
