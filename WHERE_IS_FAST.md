# Moved to WSL ext4 for speed (9p 34s → 0.6s)

Active checkout: ~/code/agent-docx (ext4)
Windows: \\wsl$\Ubuntu\home\joesa\code\agent-docx
Backup (slow): C:\Users\joesa\code\agent-docx

See ~/code/open-law-notes/WHERE_IS_FAST.md for benchmark details.

Build & run on ext4:

cd ~/code/agent-docx
pnpm install # uses /home/joesa/.pnpm-store (ext4)
pnpm run build
node ./dist/cli.js --help # ~0.65s
node ./dist/cli.js measure --help # via --help top-level

Use Node 24 for engine compliance:

~/node24/bin/node ./dist/cli.js --help # 0.53s

Original Windows checkout kept for reference; prefer ext4 for all work.
