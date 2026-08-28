# Publishing `agent-docx`

`agent-docx@0.1.0` is live on https://www.npmjs.com/package/agent-docx — `pnpm add agent-docx` works (requires Node ≥24).

## Update (bump + publish)

```sh
export PATH="$HOME/.local/node24/bin:$PATH"  # Node 24 required
cd /mnt/c/Users/joesa/code/agent-docx

# 1. bump version
npm version patch --no-git-tag-version  # or minor/major, or edit package.json manually
# or: pnpm version patch

# 2. verify pack
pnpm run build
node scripts/verify-pack.mjs          # must print "Verified agent-docx-X.Y.Z.tgz"

# 3. publish (dry-run first)
npm publish --access public --dry-run
npm publish --access public           # see Auth below

# 4. verify
npm view agent-docx version           # wait 10–30s for CDN
```

Consumers update with:

```sh
pnpm update agent-docx
agent-docx --version
agent-docx skills install --force     # if skills changed
```

## Auth

You have a granular token in `~/.npmrc`:

```
 //registry.npmjs.org/:_authToken=npm_xxxx
```

`npm whoami` → `joesalamy` means you’re set. No `npm login` needed until token expires / is revoked.

Token expired or 404/403 on `publish`:

- **With current token:** generate new one at https://www.npmjs.com/settings/tokens → `Granular Access Token` → package `agent-docx` → `Read and write` → enable `Bypass 2FA` if shown → `npm config set //registry.npmjs.org/:_authToken=npm_yyyy`
- **Without token (TOTP):** add an authenticator app at https://www.npmjs.com/settings/2fa (needs `Disable 2FA` → `Enable 2FA` → pick `Authenticator App` → scan QR → keep security key as well), then `npm publish --access public --otp=123456`
- **Security-key only:** `npm publish` won’t prompt in WSL — use the granular token path above.

Rotate immediately if token was pasted/logged: revoke at https://www.npmjs.com/settings/tokens and create a new one.

## Notes

- Node 24 lives at `~/.local/node24` — add `export PATH="$HOME/.local/node24/bin:$PATH"` to `~/.bashrc`.
- `prepack` runs `pnpm run build` — the 40s `node_modules` reinstall on `/mnt/c` is normal.
- Published `files`: `dist`, `assets`, `schemas`, `skills`, `README.md`, `LICENSE`. `verify-pack` enforces this.
