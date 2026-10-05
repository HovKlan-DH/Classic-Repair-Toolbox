# Installing CRT.Server

How to install CRT's contribution service - **CRT.Server 1.1.0, the server CRT 3.0.0 talks to** - on
the AlmaLinux box behind `classic-repair-toolbox.dk`, and how to run it afterwards. Every step ends
with a check that says plainly whether it worked.

Part one installs it from nothing, in order. Part two is for running it: deploying a new build,
reading the log, and the things you may want to change later.

> **Nothing in this file is a secret, and nothing in it should become one.** Passwords and the
> connection string are typed on the server, into a file this repository never sees. If a
> credential ever appears in a chat transcript, an email or a commit, treat it as compromised and
> change it.

## What gets installed

A small ASP.NET Core service that:

- listens on **127.0.0.1:5199 only**, and is reached through Apache at `https://classic-repair-toolbox.dk/api/`;
- runs as its own unprivileged user, `crt-server`, under systemd, restarting by itself after a crash;
- keeps its accounts, submissions and history in a MariaDB database, `crt_review`, whose tables it
  creates itself;
- writes the BETA data and the stable data (both served to CRT by Apache), the files contributors
  upload, and the feedback CRT users send;
- sends mail through the box's own postfix.

Where everything lives on this server:

```
/mydir/http/classic-repair-toolbox.dk/
├── public_html/                  <- served by Apache
│   ├── app-data-BETA/            <- the BETA source: Data/ and dataChecksums.json beside it
│   └── app-data/                 <- the stable source: Data/ and dataChecksums.json beside it
├── crt-server/                   <- NOT served by Apache
│   ├── app/                      <- the program, and appsettings.Production.json
│   ├── previous/                 <- the last build that worked, for rolling back
│   └── blobs/                    <- files contributors upload
└── user-feedback/                <- feedback attachments, opened from your network share
```

**`crt-server/` and `user-feedback/` sit BESIDE `public_html`, never inside it.** Inside, Apache would
serve `appsettings.Production.json` - database password and all - and every uploaded file to anyone
who guessed the address. Steps 5 and 6 check it.

---

# Part one - installing

## 1. The .NET runtime

The service needs the ASP.NET Core runtime - not only the base .NET runtime, and not the SDK.

```bash
sudo dnf install aspnetcore-runtime-10.0
```

**Check:**

```bash
dotnet --list-runtimes | grep -i aspnetcore      # a line with Microsoft.AspNetCore.App 10.0.x
```

Keep it patched like Apache and MariaDB: `sudo dnf upgrade` carries it.

## 2. Users and groups

```bash
sudo useradd --system --no-create-home --shell /sbin/nologin crt-server   # also makes a group crt-server
sudo groupadd -f crt-data
sudo usermod -a -G crt-data crt-server
```

Two groups, for two different jobs:

- **`crt-data`** owns the data trees and the service's folder. The service runs with this group.
- **`crt-server`** owns the feedback folder. Your network share's user may be put in it later (so you
  can delete old feedback), and that must not give it write access to the data.

**Check:** `id crt-server` lists both groups.

## 3. Folders and permissions

**The service's own error messages point here** when a folder it needs is not writable.

**The data trees.** Both must already exist, served by Apache, each holding `Data/` with the board
data and `dataChecksums.json` beside it. (For a brand-new site, copy this repository's `Assets/Data/`
into both `Data/` folders; once the service runs, Account > **Rebuild checksum manifests** in CRT's
Maintainer tab writes both `dataChecksums.json` files.) Give both trees to the `crt-data` group, with
group write and the setgid bit, so everything created in them keeps the group:

```bash
for T in /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA \
         /mydir/http/classic-repair-toolbox.dk/public_html/app-data; do
  sudo chgrp -R crt-data $T
  sudo chmod -R g+rwX    $T
  sudo find $T -type d -exec chmod g+s {} +
done
```

**The service's folder and its upload store.** The program files are owned by root, so the service
can read but never rewrite them - an attacker who got into the service could not make that stick by
editing its own files. The upload store is the one folder under it the service writes, so the
service owns that one:

```bash
S=/mydir/http/classic-repair-toolbox.dk/crt-server
sudo mkdir -p $S/app $S/blobs
sudo chown -R root:crt-data $S
sudo chmod -R 0750 $S
sudo chown -R crt-server:crt-data $S/blobs
```

**The feedback folder.** Owned by the service, group `crt-server`, setgid, so every folder and file
the service saves there gets the group and can be deleted through your share (see "Deleting feedback
through the network share" in part two):

```bash
F=/mydir/http/classic-repair-toolbox.dk/user-feedback
sudo mkdir -p $F
sudo chown crt-server:crt-server $F
sudo chmod 2775 $F
```

**Check** - each line must print its "(correct)":

```bash
for D in /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data \
         /mydir/http/classic-repair-toolbox.dk/public_html/app-data/Data \
         /mydir/http/classic-repair-toolbox.dk/crt-server/blobs \
         /mydir/http/classic-repair-toolbox.dk/user-feedback; do
  sudo -u crt-server touch $D/.probe && sudo rm -f $D/.probe && echo "$D writable (correct)" \
    || echo "*** $D NOT WRITABLE ***"
done
namei -l /mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data   # every parent traversable
```

### Copying data into the trees by hand later

**Anything you copy in as root arrives owned by root, and the service cannot write into a FOLDER made
that way.** A file is no problem - the service replaces a board file by writing a new one beside it
and renaming it into place, which needs only the folder - but a folder you copied in can only be
written by root. The service checks every folder before a publish, a publish to stable or a push-back
writes anything, and refuses with the folders named ("nothing was changed"); its log line ends with
the exact command. So after copying, run the three commands above again for that tree:

```bash
T=/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA    # or .../app-data
sudo chgrp -R crt-data $T
sudo chmod -R g+rwX    $T
sudo find $T -type d -exec chmod g+s {} +
```

The same goes for files replaced over a network share, which arrive as the share's user and lose the
group. **Do not run the service as root to avoid this.** It takes uploads from anybody on the
internet; as root, any bug in it - or in a library reading an uploaded file - would own the box.

## 4. The database

The service creates and updates its own tables when it starts. This step only makes an empty
database and a user for it:

```sql
CREATE DATABASE crt_review CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
CREATE USER 'crt_review'@'127.0.0.1' IDENTIFIED BY '<make one up on the box>';
GRANT ALL PRIVILEGES ON crt_review.* TO 'crt_review'@'127.0.0.1';
FLUSH PRIVILEGES;
```

- **`utf8mb4`, not `utf8`** - MariaDB's old `utf8` silently mangles some characters people type in
  names and comments.
- **`@'127.0.0.1'`, not `@'localhost'`.** To MariaDB those are different accounts: `localhost` is a
  socket connection, `127.0.0.1` is TCP, and the service connects over TCP.
- **The grant is for `crt_review` only**, never `*.*`.

**Check** - the `-h 127.0.0.1` is the point, it forces TCP as the service does:

```bash
mysql -u crt_review -p -h 127.0.0.1 crt_review -e "SELECT DATABASE(), CURRENT_USER();"
```

Expect `crt_review` and `crt_review@127.0.0.1`.

**One table the service writes but does not create:** `crt_update`, the launch check-ins every CRT
sends. It is older than the service and lives in `crt_review` on this server already. On a brand-new
database, create it once:

```sql
CREATE TABLE crt_review.crt_update (
  id             INT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
  createDateTime DATETIME     NOT NULL,
  ipaddr         VARCHAR(45)  NOT NULL DEFAULT '',
  versionMajor   INT          NOT NULL DEFAULT 2,
  version        VARCHAR(50)  NOT NULL DEFAULT '',
  osHighlevel    VARCHAR(100) NOT NULL DEFAULT '',
  osVersion      VARCHAR(255) NOT NULL DEFAULT '',
  cpu            VARCHAR(50)  NOT NULL DEFAULT '',
  countryCode    VARCHAR(10)  NOT NULL DEFAULT '',
  countryName    VARCHAR(100) NOT NULL DEFAULT '',
  apiJson        TEXT
) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

(When moving to a new box, take the real definition from the old one instead:
`SHOW CREATE TABLE crt_review.crt_update;`.)

## 5. Build and copy the program

On the development machine:

```bash
dotnet publish src/CRT.Server/CRT.Server.csproj -c Release -r linux-x64 --self-contained false -o ./publish-server
```

(Or Visual Studio's **Publish > Folder** with Configuration `Release`, Target Framework `net10.0`,
Target Runtime `linux-x64` and Deployment Mode **Framework-dependent** - the last two are not the
defaults. Self-contained gives hundreds of files that ignore the runtime from step 1.)

**Check what you have** before copying: about a dozen files - among them `CRT.Server.dll`,
`CRT.Data.dll`, `MySqlConnector.dll`, `CRT.Server.runtimeconfig.json`, `appsettings.Example.json` and
`crt-server.service` - plus a **`Migrations/` subfolder**. The service reads its database scripts from
`Migrations/` and will not start without it, so a copy that flattens folders breaks it.

Copy the folder to the box (to your home folder, say) and then:

```bash
APP=/mydir/http/classic-repair-toolbox.dk/crt-server/app
sudo cp -r ~/publish-server/* $APP/
sudo chown -R root:crt-data $APP
sudo chmod -R 0750 $APP
```

**The group must be `crt-data`** - the group the service runs with (`Group=` in its unit). With any
other group the `0750` leaves the service nothing, and systemd fails it with `status=200/CHDIR`.

**Check:**

```bash
ls $APP/CRT.Server.dll $APP/CRT.Data.dll $APP/crt-server.service $APP/Migrations/*.sql
sudo -u crt-server test -r $APP/CRT.Server.dll && echo "Readable (correct)"
sudo -u crt-server test -w $APP/CRT.Server.dll && echo "*** WRITABLE - fix ownership ***" || echo "Read-only to the service (correct)"

# The service folder is NOT on the web - expect 404:
curl -s -o /dev/null -w '%{http_code}\n' --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  https://classic-repair-toolbox.dk/crt-server/app/CRT.Server.dll
```

## 6. Settings

The settings live in `app/appsettings.Production.json`, beside the program - the service reads it
from there by itself. Start from the example the publish carries:

```bash
APP=/mydir/http/classic-repair-toolbox.dk/crt-server/app
sudo cp $APP/appsettings.Example.json $APP/appsettings.Production.json
sudo chown root:crt-data $APP/appsettings.Production.json
sudo chmod 0640 $APP/appsettings.Production.json
```

**`root:crt-data`, mode 0640** - the service can read its settings but not rewrite them, and nobody
else can read the password in it. **The group must be `crt-data`, not `crt-server`**: systemd gives
the service only the group its unit names, so a `root:crt-server` file is refused at start even
though `sudo -u crt-server cat` reads it fine (that command brings in every group the user is in).

**Before typing the password in, prove the file is not on the web** - expect `404`:

```bash
curl -s -o /dev/null -w '%{http_code}\n' --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  https://classic-repair-toolbox.dk/crt-server/app/appsettings.Production.json
```

Then edit it. Every setting is explained in the file itself; for this server:

| Setting | Value |
| --- | --- |
| `DataTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/Data` |
| `ManifestPath` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA/dataChecksums.json` |
| `PublicDataBaseUrl` | `https://classic-repair-toolbox.dk/app-data-BETA/Data` |
| `ProductionTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data` |
| `ProductionDataTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data/Data` |
| `ProductionManifestPath` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data/dataChecksums.json` |
| `ProductionPublicDataBaseUrl` | `https://classic-repair-toolbox.dk/app-data/Data` |
| `ProductionPublishingAdministratorsOnly` | `true` - only you publish from BETA to stable (see part two) |
| `BlobStoreRoot` | `/mydir/http/classic-repair-toolbox.dk/crt-server/blobs` |
| `FeedbackRoot` | `/mydir/http/classic-repair-toolbox.dk/user-feedback` |
| `FeedbackToAddress` | the address feedback mail goes to |
| `ConnectionString` | the database user and password from step 4 |
| `PublicApiBaseUrl` | `https://classic-repair-toolbox.dk/api` |
| `MailFromAddress` | the address the service's mail comes from |
| `AllowDataReset` | `false` (see "Resetting the contribution data" in part two) |

**A missing or wrong setting stops the service from starting**, with the setting named in the log -
it never guesses where to write. `ProductionTreeRoot` is needed either way; leave
`ProductionDataTreeRoot`, `ProductionManifestPath` and `ProductionPublicDataBaseUrl` all three empty
to run without publishing to stable (nothing then writes the stable data). If you do, also take the
stable folder out of the unit's `ReadWritePaths=` line in step 7, so the filesystem keeps it
read-only too.

**Check the service can read the file as it will really run** (same user AND group as the unit):

```bash
systemd-run --quiet --wait --pipe --uid=crt-server --gid=crt-data \
  /usr/bin/cat /mydir/http/classic-repair-toolbox.dk/crt-server/app/appsettings.Production.json > /dev/null \
  && echo "Readable by the service (correct)" || echo "*** NOT READABLE - group must be crt-data ***"
```

## 7. The systemd unit

The unit ships with the program as `crt-server.service`. Read it once - its `ReadWritePaths=` line
names the four folders from step 3, and nothing else on the box is writable to the service (without
publishing to stable, remove `.../public_html/app-data` from that line first) - then:

```bash
sudo cp /mydir/http/classic-repair-toolbox.dk/crt-server/app/crt-server.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now crt-server
```

**Check:**

```bash
systemctl status crt-server                   # active (running)
ss -ltnp | grep 5199                          # 127.0.0.1:5199 - NOT 0.0.0.0
curl -s http://127.0.0.1:5199/api/health      # {"status":"ok","version":"1.1.0",...}
journalctl -u crt-server -n 30 --no-pager | grep -v '^ '
```

On the first start the log shows the database migrations being applied, ending with
`Applied 18 migration(s).` (or however many there are); every later start says the schema is up to
date. A restart after a crash is automatic:

```bash
sudo systemctl kill -s SIGKILL crt-server; sleep 8; systemctl status crt-server   # running again
```

## 8. Apache

**SELinux first.** `getenforce` - if it says `Enforcing`, Apache may not connect to the service until
you allow it (`setsebool -P httpd_can_network_connect 1`). `Disabled` or `Permissive` (this box says
`Disabled`): nothing to do.

**The proxy.** In the `classic-repair-toolbox.dk` vhost (`/etc/httpd/conf/httpd.conf`):

```apache
ProxyPreserveHost On
ProxyPass        /api/  http://127.0.0.1:5199/api/
ProxyPassReverse /api/  http://127.0.0.1:5199/api/

# The addresses CRT 2.x and older send to, answered by the service:
ProxyPass        /app-checkin/           http://127.0.0.1:5199/api/usage/check-in
ProxyPassReverse /app-checkin/           http://127.0.0.1:5199/api/usage/check-in
ProxyPass        /app-feedback/          http://127.0.0.1:5199/api/feedback
ProxyPassReverse /app-feedback/          http://127.0.0.1:5199/api/feedback
ProxyPass        /app-contribution/api/  http://127.0.0.1:5199/api/legacy/contribution
ProxyPassReverse /app-contribution/api/  http://127.0.0.1:5199/api/legacy/contribution

RequestHeader set X-Forwarded-Proto "https" "expr=%{HTTPS} == 'on'"
```

- **The three old addresses** are where every CRT released before 3.0.0 sends its launch check-in,
  its feedback and its contributions. The service takes the check-ins and the feedback exactly as
  those CRTs send them, and tells an old CRT that tries to contribute to update.
- **Keep the `expr=` on the header**: the vhost serves both port 80 and 443, and the service builds
  the links in its mails from that header. (`ProxyPass` cannot go inside an `<If>` - Apache refuses it.)
- **Place the block before any rewrite rule that catches unknown paths**, or `/api/` answers with the
  site's own page.

**Let large feedback through.** CRT sends feedback attachments of up to 250 MB. Apache allows 1 GB
by default, but a smaller `LimitRequestBody` set anywhere would refuse it first:

```bash
grep -rin "LimitRequestBody" /etc/httpd/ 2>/dev/null   # nothing, 0, or 268435456 and up is fine
```

**Nothing in the data trees may ever run.** Contributed files land in them, so the web server must
never execute or obey anything there, whatever the service lets through. In the same vhost, once per
tree:

```apache
<Directory "/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA">
    AllowOverride None
    Options -Indexes -ExecCGI -Includes
    <FilesMatch "\.(php|phtml|phar|cgi|pl|py|sh)$">
        Require all denied
    </FilesMatch>
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

Then `apachectl configtest && sudo systemctl reload httpd`.

**Check - from the box, straight to the local interface** (testing the public address from the box
itself can hang on the router, which tells you nothing about Apache):

```bash
curl -i -m 10 --resolve classic-repair-toolbox.dk:443:127.0.0.1 https://classic-repair-toolbox.dk/api/health
# HTTP 200, "Server: Kestrel" and the JSON. HTML instead means a rewrite rule caught /api/.

curl -i -m 30 --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  -F "feedback=Test from the server" -F "version=install-test" https://classic-repair-toolbox.dk/app-feedback/
# "Success" and "Server: Kestrel", and a feedback mail a moment later.

BETA=/mydir/http/classic-repair-toolbox.dk/public_html/app-data-BETA
echo '<?php echo "RAN"; ?>' | sudo tee $BETA/Data/zz-probe.php > /dev/null
curl -s -o /dev/null -w '%{http_code}\n' --resolve classic-repair-toolbox.dk:443:127.0.0.1 \
  https://classic-repair-toolbox.dk/app-data-BETA/Data/zz-probe.php      # 403 - refused, not run
sudo rm -f $BETA/Data/zz-probe.php
```

**And from ANOTHER machine:** `curl -i https://classic-repair-toolbox.dk/api/health` answers, while
`curl -m 5 http://classic-repair-toolbox.dk:5199/api/health` must FAIL - the service is never
reachable except through Apache.

## 9. Mail

```bash
ss -ltnp | grep :25                   # postfix listening
echo "test" | sendmail -v you@example.com
```

Make sure the message actually **arrives**, and not in spam. A verification or invitation mail that
lands in spam looks exactly like a broken service.

## 10. The first administrator

Nothing in the service or in CRT can make an administrator - an endpoint that could would be one to
abuse - so the first one is made on the box. Register the account through the service, so its
password is hashed properly:

```bash
curl -s -X POST https://classic-repair-toolbox.dk/api/accounts/register -H 'Content-Type: application/json' \
  -d '{"email":"you@example.com","password":"a long passphrase","displayName":"Your name"}'
```

Open the link in the mail that arrives, then:

```sql
UPDATE crt_review.accounts SET is_administrator = 1 WHERE email_normalised = 'you@example.com';
SELECT id, email, is_verified, is_administrator, is_locked FROM crt_review.accounts
 WHERE email_normalised = 'you@example.com';
```

(`email_normalised` is the address in lower case, which is what the service matches on.) Expect one
row: verified, administrator, not locked. No restart is needed.

**Maintainers are added from CRT**, never in SQL: in CRT, tick **Enable Maintainer tab** in the
Configuration tab, sign in on the Maintainer tab, and use **Account > Maintainers** - add somebody who
has an account, or invite somebody new by email. You can add yourself too.

## 11. Check it end to end

1. In CRT's Maintainer tab, **Account > Server version** shows the version you deployed and the same
   API version for the server and for CRT.
2. Send a small change from CRT's Drafts tab. It appears under **Queue: Contributor submissions**,
   and you get a mail.
3. Approve it: it is published to BETA (`dataChecksums.json` under `app-data-BETA` changes). Check it
   in CRT with "Download data from the BETA source instead of the stable source" ticked.
4. Publish it under **Queue: Awaiting push from BETA to stable**. The stable data and its
   `dataChecksums.json` change.
5. Restart CRT: the launch check-in is a new row in `crt_review.crt_update`.

---

# Part two - running it

## Deploying a new build

Publish as in step 5, copy the folder to the box, then:

```bash
ROOT=/mydir/http/classic-repair-toolbox.dk/crt-server
sudo systemctl stop crt-server
sudo rm -rf $ROOT/previous && sudo cp -r $ROOT/app $ROOT/previous      # keep the build that worked
sudo cp -r ~/publish-server/* $ROOT/app/
sudo chown -R root:crt-data $ROOT/app && sudo chmod -R 0750 $ROOT/app
sudo systemctl start crt-server
curl -s http://127.0.0.1:5199/api/health
```

- **Stop the service before copying.** Replacing the program under a running service can crash it.
- **`appsettings.Production.json` is not in the publish**, so the copy never overwrites it.
- **The `version` in the health answer must be the new one** - compare it with the top row of
  [VERSION.md](VERSION.md). If it did not change, the copy did not land.
- **If `crt-server.service` changed** (VERSION.md says so), copy it to `/etc/systemd/system/` and
  `sudo systemctl daemon-reload` before starting. Keep any change you made to your installed one.
- A new build may need a new setting. VERSION.md says so, and if one is missing the service does not
  start and its log names it.

## Rolling back

```bash
ROOT=/mydir/http/classic-repair-toolbox.dk/crt-server
sudo systemctl stop crt-server
sudo mv $ROOT/app $ROOT/failed && sudo mv $ROOT/previous $ROOT/app
sudo systemctl start crt-server
curl -s http://127.0.0.1:5199/api/health
```

The settings file travels with `app/`, so a rollback takes its settings with it. **A rollback past a
database migration does not work**: the older build finds a migration it does not know and refuses to
start. Fix forward instead.

To stop the service entirely: `sudo systemctl disable --now crt-server`, then comment out the
`ProxyPass` lines and reload Apache.

## Is it working, and reading the log

```bash
systemctl is-active crt-server && curl -s http://127.0.0.1:5199/api/health
```

`active` and a line of JSON: it is working. If not:

```bash
journalctl -u crt-server -n 20 --no-pager -p warning     # only problems - a healthy service prints nothing
journalctl -u crt-server -n 30 --no-pager | grep -v '^ '  # recent activity, without a crash's core dump
journalctl -u crt-server -f | grep -v '^ '                # follow it live
```

A crash writes a long core dump into the journal; every line of it starts with spaces, which is what
`grep -v '^ '` drops. **A refused setting or a failed migration is not a crash**: the service logs the
reasons as critical and exits with code 78, and the unit does not restart it (a restart cannot fix
either) - `systemctl status` shows `status=78`. A database that is not up yet exits with 69 and IS
retried every five seconds.

`journalctl --vacuum-time=...` deletes history for the whole box, not one service - there is no need
to vacuum for this service.

## Troubleshooting

| What you see | Usually | How to confirm and fix |
| --- | --- | --- |
| `status=78`, the service stays stopped | A setting was refused, or a database migration failed | `journalctl -u crt-server -n 20 --no-pager -p warning` names each setting, or the migration and the database's error. Fix, then `sudo systemctl restart crt-server`. A failed migration may have left tables behind - look before dropping anything |
| "The migrations directory ... does not exist" | `Migrations/` did not reach `app/` | `ls $APP/Migrations/*.sql`; copy it across and `chown -R root:crt-data` it |
| `status=200/CHDIR` | The service cannot enter `app/` - its group is not `crt-data` | `ls -ld .../crt-server/app`; `sudo chown -R root:crt-data .../crt-server` |
| `Permission denied` reading `appsettings.Production.json` at start | The file's group is not `crt-data` | `ls -ln` it; `sudo chown root:crt-data` it. Check with the `systemd-run` line from step 6, not `sudo -u` |
| `status=150/EXEC` or "file not found" | The program was not copied, or only partly | `ls -l .../crt-server/app/` |
| `systemctl start` hangs, then fails, though the service answers | The unit says `Type=notify` | Use the shipped unit (`Type=simple`) |
| A publish, a publish to stable or a push-back is refused: "The server is not allowed to write into [...] so nothing was changed" | A folder copied in by hand, or over a share, without group write | The log line ends with the command for exactly those folders; or run the three commands in "Copying data into the trees by hand later". Then try again |
| A publish refuses "the published tree contains a symbolic link" | A link somewhere in the data tree | `find <tree> -type l`; replace each with the real folder or file |
| `/api/health` answers with the website's HTML | A rewrite rule catches `/api/` before the proxy | Move the `ProxyPass` block up |
| 503 from Apache | SELinux is enforcing and blocks the connection | `getenforce`; `setsebool -P httpd_can_network_connect 1` |
| 502 from Apache | The service is down, or the port differs | `systemctl status crt-server`; `ss -ltnp \| grep 5199` |
| The public address hangs when tried ON the box | The router hairpin, not Apache | Test with `--resolve ...:443:127.0.0.1` from the box, and the public address from another machine |
| The service answers from outside on port 5199 | `ASPNETCORE_URLS` is not loopback | Must be `http://127.0.0.1:5199`, as in the shipped unit |
| `appsettings.Production.json` downloads over the web | `crt-server/` is inside `public_html` | Move it out, then change the database password - it has been published |
| Verification, invitation or feedback mail never arrives | postfix, not the service | `journalctl -u postfix -n 30 --no-pager`; `mailq`. The service logs a warning and carries on |
| Feedback is mailed but its files are not saved | The feedback folders are full (`FeedbackMaxStoredBytes`, 20 GB) | Delete old `feedback-*` folders; the service counts again within five minutes |

## Settings you may change later

Each goes inside the `CrtServer` section of `appsettings.Production.json`; restart the service after
changing it (`sudo systemctl restart crt-server`).

- **`ProductionPublishingAdministratorsOnly`** - `true` (the default): only you publish from BETA to
  stable; maintainers see the publish greyed out with the reason, and can still push back or reject.
  `false`: maintainers publish their own systems to stable too. You are mailed each time one does.
- **`CountLocalNetworkBoardViews`** - `true` (the default) counts board views sent from your own
  network (your CRTs at home), marked `fromLocalNetwork = 1`; `false` stops storing them. Those
  already stored stay until you delete them:
  `DELETE FROM crt_review.crt_board_views WHERE fromLocalNetwork = 1;`
- **`MinimumFreeDiskBytes`** - contributions pause while the disk holding the uploads has less than
  this free (5 GB by default; `0` turns it off).
- **`FeedbackMaxStoredBytes`** - the most the saved feedback may take all together (20 GB by default;
  `0` turns it off). Past it, feedback text is still mailed but its files are not saved.
- **`RefreshTokenDays`** - how long a sign-in lasts after it was last used (30 by default). Longer
  keeps a stolen session alive longer.
- **Publishing to stable off**: empty the three `ProductionDataTreeRoot` / `ProductionManifestPath` /
  `ProductionPublicDataBaseUrl` settings. Nothing then leaves BETA, and the Maintainer tab says so.
  Take `/mydir/http/classic-repair-toolbox.dk/public_html/app-data` out of `ReadWritePaths=` in
  `/etc/systemd/system/crt-server.service` as well, then `sudo systemctl daemon-reload` before the
  restart - the filesystem then keeps the stable data read-only to the service, not only its own
  check. Put it back before switching publishing on again: with the settings set and the folder not
  writable, the service refuses to start and names the folder.

**Closing one board to contributions** (when one is being flooded) is a database switch, no restart:

```sql
UPDATE crt_review.systems SET is_accepting = 0 WHERE system_id = 'Commodore/C64/250407';   -- 1 reopens it
```

## Deleting feedback through the network share

Every folder and file the service saves under `user-feedback` is group-writable with group
`crt-server`. If deleting one through your share says "access denied", put the share's user in that
group. `testparm -s 2>/dev/null | grep -iE '^\[|path *=|force user'` shows the share whose `path`
holds `user-feedback`; its `force user` (or the name you log in to the share with) is the user. Then:

```bash
sudo usermod -a -G crt-server THATUSER
sudo systemctl restart smb      # a share picks up a new group only on a new connection
```

## Accounts and sessions

A sign-in in CRT's Maintainer tab lasts until it has not been used for `RefreshTokenDays`; every use
pushes the end further out. On Windows CRT keeps it encrypted to the Windows user; on Linux and macOS
it keeps nothing, and asks at every launch.

| To | Do |
| --- | --- |
| Sign one computer out | **Account > My account > Sign out** in CRT |
| Shut an account out at once | `UPDATE crt_review.accounts SET is_locked = 1 WHERE email_normalised = '...';` - refused on its very next request |
| Sign an account out everywhere, without locking it | `UPDATE crt_review.sessions SET revoked_utc = UTC_TIMESTAMP() WHERE account_id = ...;` |

Who is signed in, and when each session was last used:

```sql
SELECT s.id, a.email, s.created_utc, s.last_used_utc, s.expires_utc
  FROM crt_review.sessions s JOIN crt_review.accounts a ON a.id = s.account_id
 WHERE s.revoked_utc IS NULL AND s.expires_utc > UTC_TIMESTAMP()
 ORDER BY s.last_used_utc DESC;
```

## Resetting the contribution data

Account > **Reset contribution data** in CRT's Maintainer tab deletes everything people sent and did
through the service - every submission and its uploaded files, every account that is not an
administrator (so every maintainer and invitation), every system record, the history, the board
views and the API usage counts. **It never touches the BETA or stable data**, the launch check-ins or
the saved feedback; your administrator account stays. Use it once, when going live after testing:

1. **Make BETA the same as stable** - copy the stable data over the BETA data by hand (then
   "Copying data into the trees by hand later"), removing from BETA what stable does not have.
2. **Account > Rebuild checksum manifests.**
3. **Switch the reset on**: `"AllowDataReset": true` in the settings, and restart.
4. **Reset**: Account > Reset contribution data shows what goes; type `RESET` and press the button.
5. **Switch it off again**: `"AllowDataReset": false` (or remove the line), and restart.

It works only while the setting is on, so a stolen administrator session alone can never wipe the
database. Testers' CRTs then show their old submissions as "No longer on the server", and the same
drafts can be sent again.

## The database migrations

The service updates its own tables when it starts, from the numbered scripts in `app/Migrations/`,
and records each one it has applied. It **refuses to start** (`status=78`) when:

| It refuses when | Because |
| --- | --- |
| An applied script has been changed - even a comment | The database it ran against cannot change after the fact. Put the file back as it was; a change is always a NEW script |
| A number is missing | Applying a later script without an earlier one gives a database no sequence of scripts can reproduce |
| A script's number is lower than one already applied | Two builds each added "the next" script |
| A script that was applied is gone | An older build was deployed over a newer database |

**Never change the database's tables by hand.** MariaDB applies table changes as it goes, so a script
that fails halfway leaves what it already created behind: read the log, look at the database, and
clean up before restarting.

## Looking at the numbers

```bash
# The newest board views - which boards CRT users look at, and from which country:
mysql -u root -p -e "SELECT viewedUtc, hardwareName, boardName, version, countryCode, fromBeta FROM crt_review.crt_board_views ORDER BY id DESC LIMIT 10;"

# Which CRT versions called which route today (written every five minutes):
mysql -u root -p -e "SELECT callDate, method, route, version, calls FROM crt_review.crt_api_calls ORDER BY lastUtc DESC LIMIT 10;"

# The newest launch check-ins:
mysql -u root -p -e "SELECT createDateTime, version, osHighlevel, countryCode FROM crt_review.crt_update ORDER BY id DESC LIMIT 10;"
```

CRT's Maintainer tab shows the same, worded: each system's **Statistics** view, and Account > **API
usage**. No address or account is stored with a board view or an API call.
