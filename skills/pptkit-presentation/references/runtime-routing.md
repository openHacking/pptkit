# Runtime routing gate

Use this gate only after the user has approved generation. It owns the Browser-versus-Node decision and must finish before either workflow creates artifacts.

## State machine

1. Resolve the preview URL from the user or host, `PPTKIT_PREVIEW_URL`, or `https://openhacking.github.io/pptkit/`, in that order.
2. Check whether the request itself makes browser preview unsuitable:
   - an inline asset exceeds 5 MB or all inline assets exceed 20 MB;
   - the user requires unattended local output; or
   - the user requires Office/LibreOffice rendering.
3. If none applies, load the in-app Browser skill, initialize its runtime, explicitly select `iab`, and try to open the resolved URL. An abbreviated tool list is not a failed check.
4. If the page opens, continue with `browser-workflow.md`. Do not initialize a Node project.
5. If setup, selection, navigation, compatibility, or required browser API verification actually fails, record the exact failed step and returned error. Then continue with `node-workflow.md`.

Do not read or execute `node-workflow.md` while the decision is unresolved. Do not claim a browser failure without a tool result. A timeout or tool error is evidence; the absence of an initially visible control is not.

## Auditable Node decision

The guarded Node initializer requires all of these values and writes them to `runtime-decision.json` in the generated project:

- `fallback-reason`: one of `browser-setup-failed`, `preview-navigation-failed`, `preview-incompatible`, `browser-api-unavailable`, `asset-transfer-limit`, `unattended-local-output`, or `strict-office-rendering`;
- `browser-check`: `failed` for an attempted browser check or `not-required` for a measured/requested suitability condition;
- `browser-step`: `setup`, `selection`, `navigation`, `compatibility`, `api-check`, `asset-limit`, or `user-requirement`, consistent with the reason;
- `fallback-evidence`: the concrete tool result, user requirement, or measured asset size; and
- `preview-url`: the resolved HTTPS URL, defaulting to the official preview application.

The initializer rejects missing, contradictory, or vague routing evidence before it creates the output directory. Never weaken, patch around, or fabricate this receipt to make initialization proceed.
