## Cursor Cloud specific instructions

This is a Shopify embedded app ("AI Report Assistant") built on the React Router + Vite template. See `README.md` for general documentation.

### Quick reference

| Action | Command |
|--------|---------|
| Install deps | `npm install` |
| Prisma setup | `npx prisma generate && npx prisma migrate deploy` |
| Lint | `npm run lint` |
| Typecheck | `npm run typecheck` |
| Build | `npm run build` |
| Dev server (standalone) | `npx vite --port 3000` |
| Dev server (full Shopify) | `npm run dev` (requires Shopify CLI auth + partner account) |

### Running the dev server without Shopify CLI

The canonical dev command is `shopify app dev`, which handles OAuth tunneling, env var injection, and Prisma migrations. In Cloud Agent environments without Shopify CLI authentication, you can run the Vite dev server directly:

1. Ensure a `.env` file exists at the repo root with at least:
   ```
   SHOPIFY_APP_URL=http://localhost:3000
   SHOPIFY_API_KEY=test_api_key
   SHOPIFY_API_SECRET=test_api_secret
   SCOPES=read_products,read_orders,read_customers
   ```
2. Run `npx vite --port 3000`.

The landing page at `/` will render the login form. Authenticated app routes under `/app` require a real Shopify dev store session.

### Known issues

- `npm run lint` has one pre-existing error: unused variable `today` in `app/ai.server.ts`.
- The `chartjs-node-canvas` dependency requires `canvas` native binaries; these install via prebuild on Linux x64 without issues.
- SQLite database file is at `prisma/dev.sqlite` (gitignored). Prisma migrations are applied automatically by `prisma migrate deploy`.
