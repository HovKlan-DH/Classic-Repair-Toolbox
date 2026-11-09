# Installing CRT.Server

How to install CRT's contribution service - **CRT.Server 1.0.0, the server CRT 3.0.0 talks to** - on
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
from there by itself. (With automatic updates they move to `settings/`, linked into every build -
"Automatic updates from GitHub", step 4.) Start from the example the publish carries:

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
curl -s http://127.0.0.1:5199/api/health      # {"status":"ok","version":"1.0.0",...}
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
ProxyTimeout     300
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
- **`ProxyTimeout 300`**: how long Apache waits for the service's answer before giving up on it
  (Apache's own default is the vhost's `Timeout`, often 60 seconds). Publishing a whole board to BETA
  or the stable source, a push-back or deleting a board can take longer than that on a big board. The
  service finishes a write it has started whatever Apache does - a publish is never stopped half way
  because the request was dropped - but a shorter timeout tells the maintainer "the server did not
  answer" for something that worked, and CRT then has to look again to find out. 300 seconds is above
  CRT's own two-minute wait.
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

**A feedback mail can be up to 20 MB** (server 3.4.0: the pictures pasted into the text are shown in
it - `FeedbackFlow.MaxMailBytes`). Postfix refuses anything over its `message_size_limit`, 10 MB by
default, and a refused feedback reaches nobody (CRT says the server could not pass it on). Raise it
with room to spare, and check the mailbox the feedback goes to takes a mail that size too:

```bash
postconf message_size_limit                          # 10240000 = the default, too small
sudo postconf -e message_size_limit=26214400         # 25 MB
sudo systemctl reload postfix
postconf message_size_limit                          # 26214400
```

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

**Maintainers are added from CRT**, never in SQL: in CRT, tick **Enable "Maintainer" tab** in the
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

**Once "Automatic updates from GitHub" (below) is set up, the box does this itself** - and the
program then lives in `releases/<version>/`, not in `app/`, so the commands here no longer fit; that
section says how to deploy and roll back by hand then. Until it is set up:

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

(Before automatic updates are set up - after, see "Running it" under "Automatic updates from GitHub".)

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

## Automatic updates from GitHub

**Once this is set up, nobody deploys a build by hand.** Every minute the box looks on GitHub for a newer
server build, and installs it (server 10.3.0 and up):

| Instance | Installs | Built by | Unit and timer |
| --- | --- | --- | --- |
| Live service | the server of the newest **real CRT release** | the release workflow, `build-and-release.yml` | `crt-server-update.service`, `.timer` |
| Test instance | the server of **every commit pushed to main** whose tests passed | the push workflow, `build-and-unittest.yml` | `crt-server-test-update.service`, `.timer` |

So no release is needed for the test instance: a push to main is on it a few minutes later (the tests take
that long). The live service only ever gets what a real release carries. Main's builds stand on one
pre-release on GitHub, `crt-server-main`, its files replaced at every push - it is no CRT release, and CRT's
own update check never offers it.

A minute's look is cheap: it reads one small file on GitHub naming the newest build (not GitHub's API,
which allows only 60 calls an hour without a login), and downloads only a build it has not seen that
says it is newer. A commit that leaves the server's version as it was is not even downloaded.

**How it stays safe:**

- **Nothing unproven is installed, or even unpacked.** Each workflow attests the server build it makes;
  the box checks it with `gh attestation verify`, held to this repository, that instance's workflow, the
  `main` branch and a GitHub-hosted runner. A build made anywhere else, by the other workflow, or changed
  after it was built, is refused. The box needs no GitHub login and holds no key: the attestation travels
  with the build.
- **Its own user, `crt-deploy`**, which can write the builds and nothing else. It is not in
  `crt-data`, so it cannot read the settings with the database password; the service, in turn,
  still cannot write its own program. The one thing it may do beyond that is **restart** the two
  services - a polkit rule allows exactly that.
- **Never backwards, never a pre-release's build on the live service** - the build's own description,
  inside the attested bytes, says which it is - and **never a build that failed before**.
- **A build that does not come up is rolled back**: after the restart the updater asks
  `/api/health` for two minutes; without the new version answering, the build before goes back and
  the service is restarted again. **Past a database migration that cannot work** (see "Rolling
  back"): the status then says THE SERVICE IS DOWN. The test instance meets every migration first.
- **A build that ships a changed systemd unit, timer or polkit rule waits for you** - those need
  root. See "When an update waits for you" below.

What it did shows in CRT's Maintainer tab, Account > **Server version**, under the version lines.

### Setting it up, once

**The test instance first, now; the live service on CRT 4.0's release day.** The live service runs the
server CRT 3.x talks to until then, and a 4.0 server must not reach it before the release - so its
updater is set up as part of "Switching the service to the 4.0 trees", with the names at the end of this
list.

**Before starting:** push a commit to main with server 10.3.0 or newer in it, and wait for its "Build and
Unit Test" run on GitHub to finish green. Its last job, "Build the server for the test server", puts the
build on the pre-release `crt-server-main`. The box is set up from THAT build - so the very first install
already goes through the check every later one does.

Then on the box, each step in the same shell:

1. **The GitHub CLI**, from GitHub's own package repository (AlmaLinux has none of its own). It
   updates with `dnf upgrade` like everything else:

   ```bash
   sudo dnf install 'dnf-command(config-manager)'
   sudo dnf config-manager --add-repo https://cli.github.com/packages/rpm/gh-cli.repo
   sudo dnf install gh
   gh --version                                   # 2.49 or newer
   ```

   Do NOT run `gh auth login` - the updater needs no GitHub account, and should have none.

2. **The updater's user**, with a group of its own (not `crt-data`, so it cannot read the settings):

   ```bash
   sudo useradd --system --no-create-home --shell /sbin/nologin crt-deploy
   id crt-deploy                                  # groups=...(crt-deploy) - and NOT crt-data
   ```

3. **The first build, downloaded and proven:**

   ```bash
   mkdir -p ~/crt-server-first && cd ~/crt-server-first
   rm -f crt-server-linux-x64.*
   curl -fsSLO https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/download/crt-server-main/crt-server-linux-x64.tar.gz
   curl -fsSLO https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/download/crt-server-main/crt-server-linux-x64.tar.gz.sigstore.json
   gh attestation verify crt-server-linux-x64.tar.gz --bundle crt-server-linux-x64.tar.gz.sigstore.json \
     --repo HovKlan-DH/Classic-Repair-Toolbox \
     --signer-workflow HovKlan-DH/Classic-Repair-Toolbox/.github/workflows/build-and-unittest.yml \
     --source-ref refs/heads/main --deny-self-hosted-runners
   ```

   **Check:** it ends with "Verification succeeded!". Anything else: stop here. Then unpack it and read
   its version:

   ```bash
   rm -rf build && mkdir build && tar -xzf crt-server-linux-x64.tar.gz -C build
   V=$(sed -n 's/.*"serverVersion": *"\([^"]*\)".*/\1/p' build/crt-server-release.json); echo "[$V]"
   ```

   **Check:** it prints the version in brackets, `[10.3.0]` or newer.

4. **The builds side by side.** The program goes into `releases/<version>/`; `app` becomes a link to
   the build running, so the service's unit stays as it is; the settings move out of the program
   folder into `settings/`, linked into every build. The build running now is kept as
   `app-before-updates` until all is well:

   ```bash
   S=/mydir/http/classic-repair-toolbox.dk/crt-server-test
   sudo systemctl stop crt-server-test
   sudo mkdir $S/settings $S/releases
   sudo mv $S/app/appsettings.Production.json $S/settings/
   sudo chown -R root:crt-data $S/settings
   sudo chmod 0750 $S/settings
   sudo chmod 0640 $S/settings/appsettings.Production.json
   sudo mv $S/app $S/app-before-updates
   sudo cp -r ~/crt-server-first/build $S/releases/$V
   sudo ln -s $S/settings/appsettings.Production.json $S/releases/$V/appsettings.Production.json
   sudo ln -s $V $S/releases/current
   sudo ln -s releases/current $S/app
   sudo chown -hR crt-deploy:crt-data $S/releases
   sudo find $S/releases -type d -exec chmod 2750 {} +
   sudo find $S/releases -type f -exec chmod 0640 {} +
   sudo setfacl -m u:crt-deploy:x $S
   sudo cp $S/app/crt-server-test.service /etc/systemd/system/
   sudo systemctl daemon-reload
   sudo systemctl start crt-server-test
   ```

   The `2750` on the folders - the setgid bit - is what makes every file the updater unpacks later
   belong to `crt-data`, so the service can read it although the updater is not in that group. The
   `setfacl` line lets `crt-deploy` PASS THROUGH the instance's folder to its `releases/` - nothing more:
   it cannot list the folder, and `settings/` and `blobs/` stay closed to it. (`setfacl` missing: `sudo
   dnf install acl`.)

   **Check** - give it a few seconds to start:

   ```bash
   curl -s http://127.0.0.1:5198/api/health        # "version":"<V>"
   ls -l $S/app $S/releases/current                # app -> releases/current, current -> <V>
   sudo -u crt-deploy test -r $S/app/CRT.Server.dll && echo "The updater reaches its builds (correct)" \
     || echo "*** THE UPDATER CANNOT REACH ITS BUILDS - namei -l $S/app/CRT.Server.dll shows which folder ***"
   sudo -u crt-deploy cat $S/settings/appsettings.Production.json > /dev/null 2>&1 \
     && echo "*** THE UPDATER CAN READ THE SETTINGS - check step 2 ***" || echo "Settings closed to the updater (correct)"
   ```

5. **The polkit rule** that lets `crt-deploy` restart the two services and nothing else:

   ```bash
   sudo cp $S/app/50-crt-server-update.rules /etc/polkit-1/rules.d/
   sudo chmod 0644 /etc/polkit-1/rules.d/50-crt-server-update.rules
   sudo systemctl restart polkit
   cd /
   sudo -u crt-deploy systemctl --no-ask-password restart crt-server-test.service && echo "Restart allowed (correct)"
   sudo -u crt-deploy systemctl --no-ask-password stop crt-server-test.service || echo "Stop refused (correct)"
   ```

   **Check:** both lines say "(correct)".

6. **The updater's unit and timer:**

   ```bash
   sudo cp $S/app/crt-server-test-update.service $S/app/crt-server-test-update.timer /etc/systemd/system/
   sudo systemctl daemon-reload
   ```

7. **One look by hand, as a dry run** - it downloads and proves the newest build and says whether it
   would install it, installing nothing:

   ```bash
   cd /
   sudo -u crt-deploy /usr/bin/dotnet $S/app/CRT.Server.dll update-server --channel=main \
     --folder=$S --service=crt-server-test.service --health=http://127.0.0.1:5198/api/health --dry-run
   ```

   **Check:** it ends with "Up to date: <V> is the newest server build (main at [...])." A line
   with "NOT installed: GitHub's attestation does not prove ..." means the check failed - the words
   after it say why (gh not reaching Sigstore's servers reads the same); nothing was installed.

8. **Start the timer:**

   ```bash
   sudo systemctl enable --now crt-server-test-update.timer
   systemctl list-timers crt-server-test-update.timer   # its next look, within a minute
   ```

   **Check,** a couple of minutes later: `cat $S/releases/update-status.json` says `"outcome": "up-to-date"`,
   and a pre-release of CRT shows it under Maintainer > Account > **Server version**.

9. **Tidy up** once all is well: `sudo rm -rf $S/app-before-updates $S/previous ~/crt-server-first`.

**If step 4 went wrong**, back to how it was:

```bash
sudo systemctl disable --now crt-server-test-update.timer 2>/dev/null
sudo systemctl stop crt-server-test
sudo rm $S/app
sudo mv $S/app-before-updates $S/app
sudo cp $S/settings/appsettings.Production.json $S/app/
sudo chown root:crt-data $S/app/appsettings.Production.json && sudo chmod 0640 $S/app/appsettings.Production.json
sudo systemctl start crt-server-test
```

**For the live service on release day**, the same steps from step 3, from the build CRT 4.0.0's release
carries - in step 3 its downloads are `.../releases/download/4.0.0/...` and its check names
`build-and-release.yml` in place of `build-and-unittest.yml` - and with
`S=/mydir/http/classic-repair-toolbox.dk/crt-server`, port `5199`, `crt-server.service`,
`--channel=releases`, and its own unit and timer, `crt-server-update.service` and
`crt-server-update.timer`. Steps 1, 2 and the polkit rule are done already - the rule names both
services; its check in step 5 then names `crt-server.service`.

### Running it

(Below for the live service; for the test instance read `crt-server-test` and `crt-server-test-update`.)

- **What it did:** Maintainer tab, Account > **Server version**; or on the box
  `cat /mydir/http/classic-repair-toolbox.dk/crt-server/releases/update-status.json`, and its log:
  `journalctl -u crt-server-update -n 30 --no-pager`. A minute with nothing new writes nothing there -
  only what it found, installed or could not do.
- **A look now**, instead of waiting a minute: `sudo systemctl start crt-server-update.service`.
- **Builds kept:** the one running, the one it replaced and one more, in `releases/`. Older ones go
  by themselves.
- **Rolling back by hand:** point `current` at an older build and restart -
  `sudo -u crt-deploy ln -sfn <version> $S/releases/current && sudo systemctl restart crt-server`.
  The updater never installs an older build, so it leaves yours alone until a newer one comes.
- **Deploying a build by hand** is still possible: unpack it into `releases/<its version>/` as
  `crt-deploy`, link the settings into it as in step 4 of the setup, point `current` at it and restart.
- **Turning it off:** `sudo systemctl disable --now crt-server-update.timer`. The service runs on as
  it is.

### When an update waits for you, or failed

- **"... waits for you: it ships crt-server.service changed ..."** - a new build changed a unit, the
  timer or the polkit rule, which only root installs. The build is unpacked in `releases/<version>/`.
  Read VERSION.md for why, copy the named files from there to `/etc/systemd/system/` (the polkit rule
  to `/etc/polkit-1/rules.d/`) and run `sudo systemctl daemon-reload`; the next look installs the
  build. If you keep changes of your own in an installed file on purpose, merge them in, or tell the
  updater you have looked: `sudo touch $S/releases/<version>/system-files-installed`.
- **"... did not come up (failed (exit 78)), so ... runs again"** - the new build would not start:
  usually a new setting it needs (VERSION.md says so, its journal names it). Add it to
  `settings/appsettings.Production.json`, then try that build again:
  `sudo -u crt-deploy /usr/bin/dotnet $S/app/CRT.Server.dll update-server <the unit's options> --retry-failed`
  (copy the options from the unit's `ExecStart=` line). Without that, a failed build is never tried
  again by itself; the next newer one is.
- **"THE SERVICE IS DOWN"** - the new build did not start and the old one would not come back,
  almost always because the new one migrated the database. Fix forward: find why the new one does
  not start (its journal), point `current` at it again and restart.
- **"... could not be restarted ... Is the polkit rule installed?"** - step 5 of the setup. The next
  minute tries again.

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
| Account > "Server version" says mails could not be sent | postfix refused them | Fix postfix (row above). The service tries every refused mail again every few minutes for three days, then gives it up (`mail_outbox` table). A verification, password reset, address-change or invitation mail is never kept with its words (only token hashes are stored) - that person asks for a new one, or the administrator invites again |
| Feedback is mailed but its files are not saved | The feedback folders are full (`FeedbackMaxStoredBytes`, 20 GB) | Delete old `feedback-*` folders; the service counts again within five minutes |

## Settings you may change later

Each goes inside the `CrtServer` section of `appsettings.Production.json`; restart the service after
changing it (`sudo systemctl restart crt-server`).

- **`ProductionPublishingAdministratorsOnly`** - `true` (the default): only you publish from BETA to
  stable; maintainers see the publish greyed out with the reason, and can still push back or reject.
  `false`: maintainers publish their own boards to stable too. You are mailed each time one does.
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
UPDATE crt_review.boards SET is_accepting = 0 WHERE board_id = 'Commodore/C64/250407';   -- 1 reopens it
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
administrator (so every maintainer and invitation), every board record, the history, the board
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

## Converting a data tree for CRT 4.0

**Not live yet.** CRT 4.0 keeps its board data as JSON instead of Excel, and reads it from two
folders of its own. Until CRT 4.0 is released no CRT reads them and nothing but this conversion
writes them: CRT 4.0's pre-releases read the test instance's copies of them, `app-data-v4-test` and
`app-data-v4-test-BETA` ("A test instance for pre-releases", below), CRT 3.x and older keep reading
`app-data` and `app-data-BETA`, and the 3.x service keeps writing those as before. The 4.0 service
writes only the 4.0 folders - see "Switching the service to the 4.0 trees" below.

```
public_html/
├── app-data-v4-BETA/   <- CRT 4.0's BETA source: Data/ and dataChecksums.json beside it
└── app-data-v4/        <- CRT 4.0's stable source
```

Inside each, `Data/` holds the main file (`Classic-Repair-Toolbox.v4.0.0.json`), one `board.json` per
board in the board's own folder (`Commodore/C64/250407/board.json`), every file's bytes stored once
under its checksum in `Files/`, and the MiniPro IC tests.

### The two folders, once

The same group and permissions as the trees in step 3:

```bash
for T in /mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4-BETA \
         /mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4; do
  sudo mkdir -p $T
  sudo chgrp -R crt-data $T
  sudo chmod -R g+rwX    $T
  sudo find $T -type d -exec chmod g+s {} +
done
```

And in the vhost, a `<Directory>` block for each, exactly like the two in step 8 with the folder's
name changed - nothing in them may ever run either. Then `apachectl configtest && sudo systemctl
reload httpd`.

### Running the conversion

The conversion is a command of the service's own program. It runs beside the running service,
needs no settings and no database, and only READS the tree it converts. See what it would do first:

```bash
APP=/mydir/http/classic-repair-toolbox.dk/crt-server/app
WEB=/mydir/http/classic-repair-toolbox.dk/public_html
sudo -u crt-server dotnet $APP/CRT.Server.dll convert-data-tree --source=$WEB/app-data-BETA/Data --plan-only
```

It lists the boards, the files and their sizes, every file it leaves out (the older generations'
workbooks, what only they cite, and `!` notes no board cites), and - as `PROBLEM:` lines - anything
it could not carry: a column or sheet CRT does not read, a file a row cites that is missing, a board
the main workbook does not list. **With a problem it writes nothing**; fix the data and run it again.

**It also makes every board's pinouts** (since server 10.1.0): the pins its scope baselines and its
KiCad data say, and the package a footprint or the baselines make certain - the Excel trees have no
pinouts, so each conversion makes them again. A `Note:` line per board says how many components got
pins and from what; KiCad files it could not read are a note too (that board then gets its baselines'
pins alone), not a problem.

Then convert each tree into its own 4.0 folder, with that folder's public address so its manifest is
written:

```bash
sudo -u crt-server dotnet $APP/CRT.Server.dll convert-data-tree \
  --source=$WEB/app-data-BETA/Data --target=$WEB/app-data-v4-BETA \
  --public-url=https://classic-repair-toolbox.dk/app-data-v4-BETA/Data
sudo -u crt-server dotnet $APP/CRT.Server.dll convert-data-tree \
  --source=$WEB/app-data/Data --target=$WEB/app-data-v4 \
  --public-url=https://classic-repair-toolbox.dk/app-data-v4/Data
```

- **Run it as `crt-server`**, so the files belong to the service that will write these trees later.
- **Run it again whenever you like.** A second run keeps every file that is already right and removes
  whatever the new conversion no longer holds, so a folder is always exactly one conversion of its tree
  as it is now. On release day it runs once more, from the final Excel trees, just before the service
  switches to the 4.0 folders.
- **It refuses** a target holding anything else (an Excel tree, a folder of other files), a target
  inside the tree it converts or around it, and a target that already has a manifest when
  `--public-url` is missing (that manifest would describe the old files).
- **Exit code** 0: written (or planned); 1: nothing written, and why is printed; 2: the command line
  was wrong.

A plan can also be made from a PC with .NET 10, against a tree over the network share - it only reads:
`dotnet run --project src/CRT.Server -c Release -- convert-data-tree --source="Z:\mydir\http\classic-repair-toolbox.dk\public_html\app-data-BETA\Data" --plan-only`.

### Switching the service to the 4.0 trees (release day)

**Server 4.0.0 and newer publish ONLY into 4.0 trees** - a board file, the main file and stored files,
never a workbook - and tell every CRT 3.x that sends a submission or opens the Maintainer tab to update
(API revision 2). So it is deployed on CRT 4.0's release day, not before, in this order:

1. **Work the queue empty on the running (3.x) service** - approve, publish to stable or reject every
   submission waiting. A submission sent by CRT 3.x names its files by their place in the Excel tree
   ("Commodore/C64/250407/Images/sheet.png"); approved into a 4.0 tree those would become the board's
   names for them.
2. **Convert both trees one last time** (above), from the final Excel trees.
3. **Deploy the 4.0 service** and point its seven tree settings at the 4.0 folders:

| Setting | Value |
| --- | --- |
| `DataTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4-BETA/Data` |
| `ManifestPath` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4-BETA/dataChecksums.json` |
| `PublicDataBaseUrl` | `https://classic-repair-toolbox.dk/app-data-v4-BETA/Data` |
| `ProductionTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4` |
| `ProductionDataTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4/Data` |
| `ProductionManifestPath` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4/dataChecksums.json` |
| `ProductionPublicDataBaseUrl` | `https://classic-repair-toolbox.dk/app-data-v4/Data` |

4. **Install the unit that ships with 4.0.0** (step 7): its `ReadWritePaths=` line names the two
   `app-data-v4` folders and no longer `app-data` or `app-data-BETA` - the Excel trees are frozen for
   CRT 3.x from now on, and the filesystem then keeps them read-only too. `sudo systemctl
   daemon-reload && sudo systemctl restart crt-server`, and check `GET /api/health` says the new
   version and API revision 2.
5. **Its automatic updates**, from CRT 4.0.0's own server build - "Automatic updates from GitHub",
   the paragraph "For the live service on release day". From then on no server build is copied by hand.

The rest of this guide still names the Excel folders the 3.x service writes (steps 3, 6, 7 and 8);
for a 4.0 service read `app-data-v4-BETA` and `app-data-v4` there.

## A test instance for pre-releases

**Every pre-release of CRT (`...-alpha...`, `...-beta...`) talks to a test instance of the service and
downloads ITS data; every release talks to the live service and downloads the live data** - CRT decides by
its own version (owner decisions, 2026-11-09: "we need to have a new data source for the test server, so it
does not mix data with the real server"), so nothing changes in CRT on a release day. The test instance has
its own program, settings, database, port AND data trees: nothing it publishes ever reaches a CRT release,
and nothing changes in the live data when a test submission is approved. Until CRT 4.0 is released, the live
service keeps serving CRT 3.x from the Excel trees, and the test instance runs the newest build - server 4.0.0
and up. Only submissions, the Maintainer tab and the data go there: a pre-release sends its launch check-in,
its feedback and its board views to the live service, like every CRT.

|  | Live service | Test instance |
| --- | --- | --- |
| Program | `crt-server/app` | `crt-server-test/app` (-> `releases/<version>` with automatic updates) |
| Settings | `crt-server/app/appsettings.Production.json` | `crt-server-test/settings/appsettings.Production.json` |
| Port (loopback) | 5199 | 5198 |
| Address | `https://classic-repair-toolbox.dk/api/` | `https://classic-repair-toolbox.dk/crt-test/api/` |
| Database | `crt_review` | `crt_review_test` |
| Data it writes | `app-data-BETA`, `app-data` (until release day; then `app-data-v4-BETA`, `app-data-v4`) | `app-data-v4-test-BETA`, `app-data-v4-test` |
| Data CRT reads from it | every release of CRT | every alpha and beta of CRT |
| Uploads and feedback | `crt-server/blobs`, `user-feedback` | `crt-server-test/blobs`, `crt-server-test/feedback` |
| Unit | `crt-server.service` | `crt-server-test.service` |

It is the same program, the same user and groups (`crt-server`, `crt-data`) and the steps of part
one with these names. **Its accounts are its own**: a pre-release signs in to it (CRT says so above
"Sign in"), so maintainers are invited on it separately and the administrator is made in its database.
What it publishes is test data, seen by CRT's alphas and betas only.

### Setting it up, once

(For a box that has none yet. The test instance on this box was made before it had trees of its own -
"Moving it to trees of its own" below is what moved it, once.)

1. **Its data trees**, as copies of the live 4.0 trees, each manifest's addresses rewritten to the copy -
   the files and their checksums are the same, only where CRT downloads them from changes:

   ```bash
   P=/mydir/http/classic-repair-toolbox.dk/public_html
   sudo cp -a $P/app-data-v4 $P/app-data-v4-test
   sudo cp -a $P/app-data-v4-BETA $P/app-data-v4-test-BETA
   sudo sed -i 's#https://classic-repair-toolbox.dk/app-data-v4/Data/#https://classic-repair-toolbox.dk/app-data-v4-test/Data/#' $P/app-data-v4-test/dataChecksums.json
   sudo sed -i 's#https://classic-repair-toolbox.dk/app-data-v4-BETA/Data/#https://classic-repair-toolbox.dk/app-data-v4-test-BETA/Data/#' $P/app-data-v4-test-BETA/dataChecksums.json
   for T in $P/app-data-v4-test $P/app-data-v4-test-BETA; do
     sudo chgrp -R crt-data $T
     sudo chmod -R g+rwX    $T
     sudo find $T -type d -exec chmod g+s {} +
   done
   ```

   And in the vhost, a `<Directory>` block for each, exactly like the 4.0 trees' ("The two folders, once"
   above) with the folder's name changed. Then `apachectl configtest && sudo systemctl reload httpd`.

2. **Its own folders** - as step 3, the upload store owned by the service and the feedback folder
   setgid (nothing posts feedback to it, but the service needs one of its own):

   ```bash
   S=/mydir/http/classic-repair-toolbox.dk/crt-server-test
   sudo mkdir -p $S/app $S/blobs $S/feedback
   sudo chown -R root:crt-data $S && sudo chmod -R 0750 $S
   sudo chown -R crt-server:crt-data $S/blobs
   sudo chown crt-server:crt-server $S/feedback && sudo chmod 2775 $S/feedback
   ```

3. **Its database** - as step 4, with its own name and user (no `crt_update` table: check-ins go to
   the live service):

   ```sql
   CREATE DATABASE crt_review_test CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
   CREATE USER 'crt_review_test'@'127.0.0.1' IDENTIFIED BY '<make one up on the box>';
   GRANT ALL PRIVILEGES ON crt_review_test.* TO 'crt_review_test'@'127.0.0.1';
   FLUSH PRIVILEGES;
   ```

4. **The program** - publish as in step 5 and copy it into `crt-server-test/app` with the same
   `chown -R root:crt-data` and `chmod -R 0750` (then "Automatic updates from GitHub" moves it into
   `releases/`, and the settings below into `settings/`).

5. **Its settings** - as step 6, in `crt-server-test/app/appsettings.Production.json` (copied from
   `appsettings.Example.json`, `root:crt-data`, mode 0640), with these values:

   | Setting | Value |
   | --- | --- |
   | `DataTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4-test-BETA/Data` |
   | `ManifestPath` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4-test-BETA/dataChecksums.json` |
   | `PublicDataBaseUrl` | `https://classic-repair-toolbox.dk/app-data-v4-test-BETA/Data` |
   | `ProductionTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4-test` |
   | `ProductionDataTreeRoot` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4-test/Data` |
   | `ProductionManifestPath` | `/mydir/http/classic-repair-toolbox.dk/public_html/app-data-v4-test/dataChecksums.json` |
   | `ProductionPublicDataBaseUrl` | `https://classic-repair-toolbox.dk/app-data-v4-test/Data` |
   | `BlobStoreRoot` | `/mydir/http/classic-repair-toolbox.dk/crt-server-test/blobs` |
   | `FeedbackRoot` | `/mydir/http/classic-repair-toolbox.dk/crt-server-test/feedback` |
   | `ConnectionString` | `crt_review_test` and the password from 3. |
   | `PublicApiBaseUrl` | `https://classic-repair-toolbox.dk/crt-test/api` - the links in its mails |
   | `MailFromDisplayName` | `Classic Repair Toolbox (test)`, so its mails are told apart |
   | `AllowDataReset` | `true` is fine here: it wipes test accounts and submissions, never a data tree |

   The rest as for the live service (`FeedbackToAddress`, `MailFromAddress`, ...).

6. **Its unit** - the program brings `crt-server-test.service`, every line of it the live unit's but
   the five that say which instance it is (port 5198, its own folders, its own data trees):

   ```bash
   sudo cp /mydir/http/classic-repair-toolbox.dk/crt-server-test/app/crt-server-test.service /etc/systemd/system/
   sudo systemctl daemon-reload
   sudo systemctl enable --now crt-server-test
   ss -ltnp | grep 5198                          # 127.0.0.1:5198 - NOT 0.0.0.0
   curl -s http://127.0.0.1:5198/api/health      # its version, and "apiRevision":2
   ```

7. **Apache** - in the vhost, beside the `/api/` lines of step 8, then `apachectl configtest && sudo
   systemctl reload httpd`:

   ```apache
   ProxyPass        /crt-test/api/  http://127.0.0.1:5198/api/
   ProxyPassReverse /crt-test/api/  http://127.0.0.1:5198/api/
   ```

   **Check:** `curl -i -m 10 --resolve classic-repair-toolbox.dk:443:127.0.0.1
   https://classic-repair-toolbox.dk/crt-test/api/health` - HTTP 200, `Server: Kestrel`, its JSON.

8. **Its first administrator** - as step 10, against its address and database:

   ```bash
   curl -s -X POST https://classic-repair-toolbox.dk/crt-test/api/accounts/register -H 'Content-Type: application/json' \
     -d '{"email":"you@example.com","password":"a long passphrase","displayName":"Your name"}'
   ```

   Open the link in the mail ("Classic Repair Toolbox (test)"), then
   `UPDATE crt_review_test.accounts SET is_administrator = 1 WHERE email_normalised = 'you@example.com';`

9. **End to end, with a CRT pre-release**: the Maintainer tab says it signs in to the test server;
   Account > **Server version** shows the test instance's version and API 2 on both lines; CRT's log names
   `app-data-v4-test` for its data; a small change sent from the Contribute tab (whose Submit dialog says it
   goes to the test server) reaches the queue; approving it changes `app-data-v4-test-BETA/dataChecksums.json`,
   publishing it to stable `app-data-v4-test/dataChecksums.json` - and nothing under `app-data-v4`.

### Moving it to trees of its own (once, 2026-11-09)

This box's test instance published into the 4.0 trees until server 10.3.1. To move it - server 10.3.1's
unit names the new trees, so its updater HOLDS 10.3.1 until that unit is installed, and installs it by
itself the minute after:

1. Wait until `cat .../crt-server-test/releases/update-status.json` says `"outcome": "held"` for 10.3.1 -
   the build is then in `releases/10.3.1/`.
2. Stop the test instance, and make its trees as in step 1 of "Setting it up, once" (the copies, the
   manifests rewritten, the groups, the two `<Directory>` blocks).
3. Point its seven tree settings at them (a copy of the file kept beside it first):

   ```bash
   F=/mydir/http/classic-repair-toolbox.dk/crt-server-test/settings/appsettings.Production.json
   sudo cp -p $F $F.before-test-trees
   sudo sed -i \
     -e 's#public_html/app-data-v4-BETA/#public_html/app-data-v4-test-BETA/#g' \
     -e 's#public_html/app-data-v4/#public_html/app-data-v4-test/#g' \
     -e 's#public_html/app-data-v4"#public_html/app-data-v4-test"#g' \
     -e 's#\.dk/app-data-v4-BETA/Data#.dk/app-data-v4-test-BETA/Data#g' \
     -e 's#\.dk/app-data-v4/Data#.dk/app-data-v4-test/Data#g' \
     $F
   sudo grep -o 'app-data[^"]*' $F               # seven lines, every one app-data-v4-test...
   ```

4. Install 10.3.1's unit and start the test instance - still the build before, now on its own trees:
   `sudo cp .../crt-server-test/releases/10.3.1/crt-server-test.service /etc/systemd/system/ && sudo
   systemctl daemon-reload && sudo systemctl start crt-server-test`.
5. The next minute the updater installs 10.3.1 and restarts it: `update-status.json` says `"installed"`.

### Running it

- **New builds:** with "Automatic updates from GitHub" set up for it, it installs the server of every
  commit pushed to main by itself, minutes after the push. By hand: as "Deploying a new build" in part two, with
  `ROOT=/mydir/http/classic-repair-toolbox.dk/crt-server-test`, `crt-server-test` and port 5198.
- **Its log:** the commands in "Is it working, and reading the log" with `-u crt-server-test`.
- **Starting over:** "Resetting the contribution data" works here as on the live service (with
  `AllowDataReset` true).
- **Its data brought up to date with the live data:** stop it, delete its two trees, make them again as in
  step 1 of "Setting it up, once", and start it. What it had published is gone with them.

### On release day

**Nothing for the test instance.** It never writes the 4.0 trees that become the live data that day, and
CRT's alphas and betas go on reading its own. Bring its data up to date with the freshly converted live
data afterwards if wanted ("Running it").

**CRT's side needs nothing either:** a release talks to the live service and reads the live data, a
pre-release the test instance's, by its version (`ContributionServer.RootUrlFor`, `ChecksumsUrlFor`).

## Looking at the numbers

```bash
# The newest board views - which boards CRT users look at, and from which country:
mysql -u root -p -e "SELECT viewedUtc, hardwareName, boardName, version, countryCode, fromBeta FROM crt_review.crt_board_views ORDER BY id DESC LIMIT 10;"

# Which CRT versions called which route today (written every five minutes):
mysql -u root -p -e "SELECT callDate, method, route, version, calls FROM crt_review.crt_api_calls ORDER BY lastUtc DESC LIMIT 10;"

# The newest launch check-ins:
mysql -u root -p -e "SELECT createDateTime, version, osHighlevel, countryCode FROM crt_review.crt_update ORDER BY id DESC LIMIT 10;"
```

CRT's Maintainer tab shows the same, worded: each board's **Statistics** view, and Account > **API
usage**. No address or account is stored with a board view or an API call.
