# Contributing

Thanks for your interest in contributing to PPTKit.

## Before You Start

- Check existing issues and discussions before opening a new one.
- For larger architectural or API changes, update the owning documentation first so the design and implementation stay aligned.
- Keep changes focused and easy to review.

## Local Development

The monorepo workspace is now bootstrapped. The standard validation workflow is:

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
