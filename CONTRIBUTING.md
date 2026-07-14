# Contributing

Thanks for your interest in contributing to PPTKit.

## Before You Start

- Check existing issues and discussions before opening a new one.
- For larger architectural or API changes, update the owning documentation first so the design and implementation stay aligned.
- Keep changes focused and easy to review.

## Local Development

The complete monorepo setup and validation workflow is documented in the
[Developer Workflow](docs/guides/developer-workflow.md). The standard validation
commands are:

```bash
pnpm install
pnpm lint
pnpm typecheck
pnpm test
```

To compile all initial packages:

```bash
pnpm build
```

## Publishing a preview release

Publishing is a local, interactive operation. First authenticate with npm and verify
that the account can publish the `@pptkit` scope:

```bash
npm login
npm whoami
```

Run a complete non-publishing check with:

```bash
pnpm release:npm --dry-run
```

When the release gate is satisfied, run `pnpm release:npm`. It publishes all public
packages under `packages/` at one unified version after lint, typecheck, tests,
build, package-content checks, and confirmation. The command does not create Git
commits or tags; review and commit the version changes manually after publishing.

See [Release Strategy](docs/migration/release-strategy.md) for failure handling and
the complete release policy.

## Contribution Workflow

1. Fork the repository.
2. Create a feature branch.
3. Make your change with tests or documentation updates when appropriate.
4. Submit a pull request with a clear description of the problem and solution.

## Scope Expectations

- Prefer small, composable changes over large rewrites.
- Keep package boundaries explicit.
- Add or update docs when behavior or architecture changes.
- Raise design questions early if a change affects public APIs.

## Coding Guidelines

- Favor clear interfaces over clever abstractions.
- Keep naming consistent across packages and docs.
- Avoid introducing broad plugin surfaces before the core architecture is stable.

## Documentation Guidelines

- Update the relevant guide, concept, or architecture doc alongside code changes.
- Prefer examples that are easy to copy, run, and verify.
- Mark provisional APIs clearly until preview publishing begins.

## Need Help?

Open a discussion or draft PR if you want early feedback.
