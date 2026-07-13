# PPTKit Contribution Guidelines

This repository is a TypeScript monorepo. Keep changes easy to understand, test, and review.

## General design rules

- Organize code by cohesive responsibility; avoid catch-all modules and element-type-specific package fragmentation.
- Keep public APIs small and intentional. Put implementation details behind package entry points.
- Prefer explicit data flow and stable contracts over hidden global state or implicit coupling.
- Keep pure transformations separate from filesystem, network, process, and other environment-dependent side effects.
- Preserve existing behavior and public API compatibility unless a breaking change is explicitly required.
- Do not introduce a new abstraction, package, or dependency unless it owns a clear responsibility and removes meaningful coupling.

## Dependency and package rules

- Dependencies must follow the package architecture documented in `docs/architecture/`.
- Do not create circular dependencies.
- Import other packages through their public exports; do not depend on `dist`, generated files, or private implementation paths.
- Keep package entry points focused on public exports and orchestration rather than low-level domain logic.
- Do not leak implementation-specific concerns into lower-level or shared packages.

## Change and testing rules

- Before editing, inspect the existing implementation, tests, and relevant architecture documentation.
- Keep unrelated working-tree changes intact.
- Every behavior change should have focused test coverage at the narrowest useful level, plus package/integration coverage where boundaries are affected.
- Update public API or architecture documentation when behavior or contracts change.
- Run the relevant type checks, build, lint, and tests before delivery.
- Do a final adversarial review: check compatibility, failure behavior, unnecessary complexity, and likely regressions.

## Review checklist

- Is the change within the owning package or module's responsibility?
- Does it introduce an invalid dependency direction or hidden coupling?
- Are side effects isolated and failure modes explicit?
- Is the public contract still intentional and compatible?
- Are tests and documentation sufficient for the changed behavior?
