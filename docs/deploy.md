# Autodeploy

Autodeploy runs from GitHub Actions after the `CI` workflow succeeds on `master`.
Secrets, client files and generated documents are not stored in Git.

## Server prerequisites

Install on the server:

- Git
- Docker
- Docker Compose plugin

Clone the repository once:

```bash
git clone https://github.com/EgorArutiunean/pretenzia.git /opt/pretenzia
cd /opt/pretenzia
```

Create production config on the server only:

```bash
cp .env.example .env
nano .env
```

Put the address dictionary into persistent storage as `storage/object_addresses.xlsx`.
For Docker/Coolify set `OBJECT_ADDRESSES_PATH=/app/storage/object_addresses.xlsx`.

Run once manually to verify the server:

```bash
docker compose build
docker compose up -d
docker compose logs -f bot
```

## GitHub secrets

Create repository secrets in GitHub:

- `DEPLOY_HOST` - server IP or DNS name.
- `DEPLOY_USER` - SSH user.
- `DEPLOY_SSH_KEY` - private SSH key allowed to connect to the server.
- `DEPLOY_PATH` - project path on the server, for example `/opt/pretenzia`.
- `DEPLOY_PORT` - optional SSH port, defaults to `22`.
- `DEPLOY_KNOWN_HOSTS` - optional pinned server host key output.

The Telegram token, admin ids, `.env`, source Excel files, address dictionary, generated ZIP files and `storage/` stay only on the server.

## Deploy flow

On every successful CI run on `master`, GitHub Actions connects to the server and runs:

```bash
cd "$DEPLOY_PATH"
git fetch origin master
git checkout master
git pull --ff-only origin master
docker compose build
docker compose up -d --remove-orphans
docker compose ps
```

Manual deploy is available from GitHub Actions -> Deploy -> Run workflow.
