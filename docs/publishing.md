# Publishing to npm

Only `@byfungsi/nori` is published. The root, demo, and internal packages remain private; the facade bundles internal code and exposes stable subpaths. Version `0.1.0` is published on npm.

## Prepare an archive

Use Node 22.20.0 and install dependencies with `npm ci`. Then run:

```sh
npm run release:prepare
```

This runs boundaries, type checks, tests, builds, and isolated consumer checks, packs the facade, checks every exported file, and dry-runs publication. The archive and its file list/integrity are saved under `artifacts/release/`. Nothing is uploaded to npm. React stays an optional peer dependency.

Nori uses the MIT license. Its text is included in the repository and the published package. The publishing workflow checks for the license metadata and bundled license file.

## First release

An npm owner of the `byfungsi` organization can bootstrap the package locally:

```sh
npm whoami
npm org ls byfungsi
npm run release:prepare -- --require-license
npm publish ./artifacts/release/byfungsi-nori-0.1.0.tgz --access public --registry https://registry.npmjs.org/
```

The last command publishes publicly and cannot overwrite an existing version. Run it only when ready to release. npm may request account verification or a one-time password. Keep credentials out of repository files. Initial local publication does not have GitHub Actions provenance.

Verify the published version and test installation in a separate project:

```sh
npm view @byfungsi/nori@0.1.0 version dist.integrity
npm install @byfungsi/nori@0.1.0
```

Compare the registry integrity to `artifacts/release/manifest.json`. Keep the installation documentation aligned with each release.

## Subsequent releases from GitHub

Configure a trusted publisher in the npm package settings:

| Setting           | Value                                 |
| ----------------- | ------------------------------------- |
| Provider          | GitHub Actions                        |
| Organization      | `byfungsi`                            |
| Repository        | `nori`                                |
| Workflow filename | `npm-publish.yml`                     |
| Environment       | `npm`                                 |
| Allowed action    | Direct publication with `npm publish` |

Create the matching GitHub environment `npm`; restrict it to the `main` branch. Required reviewers can be added if desired. The workflow uses short-lived OIDC credentials, with no stored npm publishing token. npm generates provenance for supported trusted publication from public repositories. See [npm trusted publishing](https://docs.npmjs.com/trusted-publishers/) for setup requirements.

For each release:

1. Update the public package version with `npm version 0.1.1 --workspace @byfungsi/nori --no-git-tag-version` (substitute the intended version), and commit the manifest and lockfile alongside release notes.
2. Push the reviewed changes to `main` and let CI pass.
3. Open GitHub Actions → **Publish npm package** → **Run workflow**. Select `main` and enter the exact package version.
4. The workflow checks the version and license, runs all verification, saves the archive as an Actions artifact, and publishes that exact archive.
5. Verify the registry version/integrity and create a matching `v0.1.1` Git tag on the workflow's commit for traceability.

Ordinary pushes and tags never trigger npm publication. A duplicate version fails; release fixes under a new version. The workflow publishes to `latest`; prerelease channels require an explicit workflow change before use.
