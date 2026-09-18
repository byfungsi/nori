# Contributing and verification

## Local development

```sh
npm ci
npm run dev          # React demo
npm run docs:dev     # Documentation, port 4179
```

Package source lives in `packages/{model,xlsx,core,formula,pivot,react,nori}`. The facade is the only intended consumer package. `examples/react-demo` is the browser host, while `docs/` is the VitePress documentation application.

## Required verification

```sh
npm run verify       # Boundaries, typechecks, tests, package/demo builds, packed consumer
npm run docs:build   # Library declarations, agent exports, static docs, output checks
```

For browser checks, start the demo then run `npm run verify:browser`. To verify the documentation UI, build it, run `npm run docs:preview`, then run `npm run docs:verify-browser`. Set `NORI_BROWSER_EXECUTABLE` if agent-browser needs an explicit Chrome executable path. These commands use a locally installed agent-browser CLI; it is not needed to build the library or static documentation.

## Documentation ownership

Edit Markdown in `docs/`. Keep implementation status in `support.md` distinct from future intentions in `roadmap.md`. Add new pages to the sidebar in `.vitepress/config.mts`. Complete examples live in `docs/examples`; the quick start is executed by `tests/docs.test.ts`, and examples are included in the workspace typecheck.

The generator builds plain-text and agent artifacts into `docs/public`. Do not hand-edit generated files. `docs:build` refreshes package declarations first, expands code includes, records source/content hashes, and builds the site with dead-link validation. The generated directory, VitePress cache and build output are ignored by version control.

## Hosting

`npm run docs:build` produces a static site at `docs/.vitepress/dist`. A root-hosted local build is the default. For the published GitHub Pages site, use:

```sh
NORI_DOCS_BASE=/nori/ npm run docs:build
```

The base path is applied to site assets, navigation, the agent manifest and plain-text download links. Project-site builds use explicit `.html` page routes so direct links do not depend on server rewrites.

The `pages.yml` workflow verifies the library, builds the docs, uploads the static artifact, and deploys to GitHub Pages on pushes to `main` or manual dispatch. Pull requests run verification without deployment. Pages uses the GitHub Actions source, `pages: write` and `id-token: write` deployment permissions, and the `github-pages` environment. No custom deployment credential is needed.

The published site is [byfungsi.github.io/nori](https://byfungsi.github.io/nori/). All generated text/Markdown/type artifacts are served under the same `/nori/` prefix. The npm package is not automatically published by this workflow.
