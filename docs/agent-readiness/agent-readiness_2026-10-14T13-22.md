# Agent Readiness Evaluation — `sheetrs`

Based on [Factory AI Agent Readiness](https://docs.factory.ai/web/agent-readiness/overview),
adapted for modern tooling standards and agent platform conventions.

## Ecosystem

| Property | Value |
| --- | --- |
| Date | 2026-10-14T13:22 |
| Repo type | `both` |
| Project type | `cli-tool`, `library`, `web-app` |
| Language (programming) | Rust |
| Language (docs) | English |
| Framework | None (custom parsers) |
| Stack | Rust 1.99 (edition 2024) + quick-xml + zip + clap + serde + wasm-bindgen |
| Package manager | Cargo |
| Detected tools | cargo fmt, cargo clippy, cargo test, pre-commit, just, Dependabot, GitHub Actions |

> [!NOTE]
> Monorepo with three project types. A criterion is N/A only if N/A for **all**
> declared types. The only N/A criterion is O4 (Distributed tracing).
> D3 is scored as two criteria (D3.1-D3.2 at L1, D3.3 at L2), so the
> applicable total is 53.

## Current Readiness Summary

| Pillar | Estimated Level | Score | Key Gap |
| --- | --- | --- | --- |
| Style & Validation | 🟢 **L3** | 6 / 7 | CI and pre-commit lint the whole workspace (S7) |
| Build System | 🟡 **L2** | 3.5 / 5 | CI does not pass `--locked`; no Dockerfile |
| Testing | 🟢 **L3** | 7 / 9 | No coverage in CI, no threshold |
| Documentation | 🟡 **L2** | 8 / 11 | No versioned agent skills; API docs not published |
| Dev Environment | 🔴 **L1** | 1 / 5 | No devcontainer, Docker, env template |
| Debugging & Observability | 🔴 **L1** | 0.5 / 4 | No logging or error tracking |
| Security | 🟢 **L3** | 4 / 5 | `cargo audit` only local; `main` lacks required checks |
| Task Discovery | 🟢 **L3** | 3 / 4 | No project board |
| Product & Experimentation | 🔴 **L1** | 0 / 3 | No analytics or feature flags |

**Overall: L2 — Documented (33 / 53 applicable = 62.3%)**

| Level set | Score | Progress | Threshold met |
| --- | --- | --- | --- |
| L1 → L2 | 12 / 13 | 92.3% | ✅ |
| L2 → L3 | 4.5 / 10 | 45% | ❌ |
| L3 → L4 | 6.5 / 10 | 65% | ❌ |
| L4 → L5 | 10 / 20 | 50% | ❌ |

> Criteria marked **N/A** are excluded from the scoring denominator based on
> the project type (`cli-tool`, `library`, `web-app`) and repository type
> (`both`).

> [!IMPORTANT]
> Previous evaluation (2026-05-03): L2 — Documented (30 / 52 = 57.7%).
> The CI workflow (`.github/workflows/ci.yml`) unlocks S5, T4 and half of B4.
> The migration from `AGENTS.md` + repo skills to `CLAUDE.md` removed the
> versioned agent skills (D3.3: ✅ → ❌). Branch rulesets and Dependabot
> security updates were configured in the repository settings during this
> evaluation.

---

## Detailed Pillar Evaluation

### 1. Style & Validation (6 / 7)

| Criterion | Status | Evidence |
| --- | --- | --- |
| S1 Linter configured | ✅ | `cargo clippy`; workspace lints in `Cargo.toml` |
| S2 Code formatter | ✅ | `rustfmt` |
| S3 Type checker | ✅ | Rust is statically typed |
| S4 Pre-commit hooks | ✅ | `.pre-commit-config.yaml` with `cargo fmt` and `cargo clippy` hooks |
| S5 CI enforces lint/format | ✅ | `ci.yml` jobs `fmt` (`cargo fmt --all -- --check`) and `clippy` (`-D warnings`) |
| S6 Import sorting | ✅ | `rustfmt` sorts imports |
| S7 Linter targets changed code | ❌ | Both hooks use `pass_filenames: false` and CI runs `cargo clippy --workspace`: linting covers the whole workspace |

> [!NOTE]
> Clippy and rustc check whole crates, so file-level scoping is not available
> for Rust. The rubric still scores workspace-wide linting as ❌; the
> previous evaluation scored it ✅.

### 2. Build System (3.5 / 5)

| Criterion | Status | Evidence |
| --- | --- | --- |
| B1 Build command documented | ✅ | README: `cargo build`, `cargo install --path ...` |
| B2 Dependencies pinned | ✅ | `Cargo.lock` committed |
| B3 Single-command install | ✅ | `cargo build` builds the workspace |
| B4 Reproducible CI builds | ⚠️ | CI resolves from the committed `Cargo.lock`, but without `--locked` Cargo may rewrite an outdated lockfile instead of failing |
| B5.1 Dockerfile exists | ❌ | None |
| B5.2 Multi-stage build | ❌ | — |
| B5.3 `.dockerignore` | ❌ | — |
| B5.4 Non-root user | ❌ | — |
| B5.5 Pinned base image | ❌ | — |

**B5 Score**: 0 / 5 → ❌

### 3. Testing (7 / 9)

| Criterion | Status | Evidence |
| --- | --- | --- |
| T1 Unit tests exist | ✅ | `#[test]` in 57 files across the workspace crates |
| T2 Test framework configured | ✅ | `cargo test`; `tempfile` dev-dependency |
| T3 Tests runnable locally | ✅ | README and CONTRIBUTING: `cargo test`; `just test` |
| T4 CI runs tests | ✅ | `ci.yml` job `test`: `cargo test --workspace` |
| T5 Coverage tracking | ⚠️ | `just coverage` (cargo-tarpaulin), local and optional |
| T6 Coverage threshold | ❌ | None |
| T7 Integration tests | ✅ | `sheetrs/tests/` (writer, rule dependencies, output order, docs sync) |
| T8 Test factories/fixtures | ⚠️ | Ad-hoc helpers and `tests/minimal_test.{xlsx,ods}`; no fixture framework |
| T9 TDD methodology | ✅ | See below |

**T9 — TDD Methodology**

| Check | Status | Evidence |
| --- | --- | --- |
| TDD documented in contributing/coding patterns | ✅ | CONTRIBUTING "Test Patterns": positive and negative cases required |
| Positive test cases present | ✅ | Across rule, reader and writer modules |
| Negative test cases present | ✅ | No-violation assertions (`violations.is_empty()`) and error paths (`is_err()`, `should_panic`) |

**T9 Score**: ✅

### 4. Documentation (8 / 11)

| Criterion | Status | Evidence |
| --- | --- | --- |
| D1 README with setup | ✅ | See D1 below |
| D2 Architecture docs | ✅ | `ARCHITECTURE.md`, `docs/README.md` index, `docs/coding-patterns.md` |
| D3.1-D3.2 Agent context | ✅ | See D3 below |
| D3.3 Agent skills | ❌ | See D3 below |
| D4 Env vars documented | ✅ | README: no environment variables; configuration via `sheetlint.toml` |
| D5 API documentation | ⚠️ | Rule reference in `sheetlint/README.md`; no published API docs |
| D5.1 User manual | ⚠️ | Per-tool READMEs, no consolidated manual |
| D5.2 Auto-generated docs | ⚠️ | `just doc` (`cargo doc`), not published |
| D6 Contributing guide | ✅ | `CONTRIBUTING.md`: setup, parity, doc-sync table, fragments, PR process |
| D7.1 Changelog exists | ✅ | `CHANGELOG.md` (Keep a Changelog) |
| D7.2 Progressive fragments | ✅ | `changelog.d/`, one fragment per PR, documented in CONTRIBUTING |
| D8 Code conventions | ⚠️ | See D8 below |

**D7 Score**: ✅

**D1 — README Content Quality**

| Check | Status |
| --- | --- |
| Title + one-line description | ✅ |
| Tech stack + version | ✅ (Rust ≥ 1.99, edition 2024; matches `rust-toolchain.toml`) |
| Prerequisites | ✅ |
| Build/test/lint/fmt commands | ✅ |
| Project structure | ✅ |
| Environment variables | ✅ |
| Accuracy | ✅ |

**D2 — Docs Content Quality**: `docs/README.md` indexes architecture, coding
patterns, profiling, benchmarks, test assets and the WASM demo. All links
resolve.

**D3 — Agent Context**

| Sub-criterion | Status | Evidence | Source |
| --- | --- | --- | --- |
| D3.1 Agent workflow references | ✅ | Load-before-working list (`docs/coding-patterns.md`, `CONTRIBUTING.md`), on-demand docs, benchmark preparation | `CLAUDE.md` |
| D3.2 Doc maintenance rules | ✅ | Points to CONTRIBUTING's doc-sync table; benchmark docs update rule | `CLAUDE.md` |
| D3.3 Minimal agent skills | ❌ | No `.claude/skills/`, `.claude/commands/` or `.claude/agents/`; `CLAUDE.md` relies on externally installed skills | — |

**D3.1-D3.2 Score** (L1): ✅ — both present and accurate.

**D3.3 Score** (L2): ❌ — the skills are not versioned in the repo, and no
`.claude/settings.json` declares the plugin, so a fresh clone has neither a
review nor a documentation skill.

**Agent-specific vs general check**: `CLAUDE.md` holds only loading order,
benchmark preparation and the agent-readiness re-run trigger. No general
restrictions are misplaced.

**Accuracy notes**: All file references in `CLAUDE.md` exist
(`docs/coding-patterns.md`, `CONTRIBUTING.md`, `README.md`, `ARCHITECTURE.md`,
`docs/README.md`, `scripts/README.md`, `sheetlint/docs/PROFILING.md`).

**D8 — Domain Convention Completeness**

| Check | Status | Evidence |
| --- | --- | --- |
| General conventions (style, naming, patterns) | ✅ | SoC, Rule/Walker patterns, formula normalization, configuration parameters, performance guidelines |
| `cli-tool` conventions | ✅ | Exit codes, output streams, output formats |
| `library` conventions | ✅ | Versioning, public API surface, backward compatibility |
| `web-app` conventions | ❌ | No accessibility or component conventions for the WASM demo |

**D8 Score**: ⚠️

### 5. Dev Environment (1 / 5)

| Criterion | Status | Evidence |
| --- | --- | --- |
| E1 Environment template | ❌ | No `.env.example` (no environment variables exist) |
| E2 Devcontainer | ❌ | No `.devcontainer/` |
| E3 Docker Compose local | ❌ | None (applies via `web-app`) |
| E4 One-command setup | ✅ | `justfile`: `just check` runs fmt, clippy and tests |
| E5 Seed data mechanism | ❌ | Only test assets in `tests/` |

### 6. Debugging & Observability (0.5 / 4)

| Criterion | Status | Evidence |
| --- | --- | --- |
| O1 Structured logging | ❌ | No `log`/`tracing` dependency |
| O2 Error tracking | ❌ | None (applies via `web-app`) |
| O3 Debug tools documented | ⚠️ | `sheetlint/docs/PROFILING.md`; no debugger instructions |
| O4 Distributed tracing | N/A | N/A for all declared types |
| O5 Health check | ❌ | None (applies via `web-app`) |

### 7. Security (4 / 5)

| Criterion | Status | Evidence |
| --- | --- | --- |
| X1 CODEOWNERS | ✅ | `.github/CODEOWNERS`: `* @cosmoscalibur` |
| X2 Security linting | ⚠️ | `just audit` (`cargo audit`), local and optional; not in CI |
| X3 Dependency scanning | ✅ | Dependabot for `cargo` and `github-actions`, grouped, native cooldown 15–45 days |
| X4 Secret scanning | ✅ | GitHub secret scanning and push protection enabled (repository settings) |
| X5 Branch protection | ⚠️ | Ruleset `develop`: PR required, the four CI jobs required, no deletion or force-push. Ruleset `main`: no deletion or force-push only. No approving review required (single maintainer) |

**X3 note**: Vulnerability alerts and Dependabot security updates are
enabled, so a vulnerable dependency gets a PR outside the monthly,
cooled-down schedule.

### 8. Task Discovery (3 / 4)

| Criterion | Status | Evidence |
| --- | --- | --- |
| K1 Issue templates | ✅ | `bug_report.yml`, `feature_request.yml` |
| K2 PR template | ✅ | `.github/pull_request_template.md` |
| K3 Issue labeling | ✅ | `.github/labels.yml` |
| K4 Project board | ❌ | No GitHub Project linked |

**K1 — Issue Template Content**

| Check | Status |
| --- | --- |
| Bug report template (structured) | ✅ (description, steps, expected/actual, tool, format, version, OS, reproduction file) |
| Feature request template (structured) | ✅ (problem, proposed solution, alternatives, tool, rule category) |

**K2 — PR Template Content**

| Sub-criterion | Status | Evidence |
| --- | --- | --- |
| K2.1 Issue link placeholder | ✅ | `Closes #` |
| K2.2 Description + checklist + changelog reminder | ✅ | Description, changes, testing checklist with parity, changelog fragment |

**K2 Score**: ✅

**K3 — Label Consistency**: types (`bug`, `feature`, `enhancement`,
`documentation`, `breaking-change`), rule categories, formats and PR status
(`ready-for-review`, `needs-tests`) cover both issue and PR workflows. No
mismatches for `both`.

### 9. Product & Experimentation (0 / 3)

| Criterion | Status | Evidence |
| --- | --- | --- |
| P1 Analytics instrumentation | ❌ | None in the WASM demo |
| P2 Feature flags | ❌ | Cargo features are build-time only |
| P3 Experiment infrastructure | ❌ | None |

---

## Score Comparison

| Pillar | 2026-05-03 | 2026-10-14 | Δ |
| --- | --- | --- | --- |
| Style & Validation | 6 / 7 | 6 / 7 | — (S5 ✅, S7 ❌) |
| Build System | 3 / 5 | 3.5 / 5 | +0.5 |
| Testing | 5.5 / 9 | 7 / 9 | +1.5 |
| Documentation | 8.5 / 10 | 8 / 11 | −0.5 |
| Dev Environment | 1 / 5 | 1 / 5 | — |
| Debugging & Observability | 0.5 / 4 | 0.5 / 4 | — |
| Security | 2.5 / 5 | 4 / 5 | +1.5 |
| Task Discovery | 3 / 4 | 3 / 4 | — |
| Product & Experimentation | 0 / 3 | 0 / 3 | — |
| **Total** | **30 / 52** | **33 / 53** | **+3** |
| **Percentage** | **57.7%** | **62.3%** | **+4.6pp** |
| **Level** | **L2** | **L2** | — |

---

## Modernization Recommendations

No legacy tools detected.

| Current | Recommended | Rationale | Breaking? |
| --- | --- | --- | --- |
| `cargo audit` via `just audit` | `cargo audit` (or `cargo-deny`) job in CI | Enforce X2 instead of relying on a local recipe | No |
| `cargo-tarpaulin` via `just coverage` | `cargo-llvm-cov` in CI with `--fail-under-lines` | Coverage in CI with a threshold (T5, T6) | No |

---

## Proposed Enhancement Plan

> [!IMPORTANT]
> This plan addresses **configuration and documentation changes only** unless a
> modernization migration explicitly requires code changes.

### Phase 1 — Unlock Level 3 (L2 criteria)

> **Projected level after Phase 1: L2 (L2 set 7 / 10 = 70%)**

#### [NEW] .claude/settings.json or .claude/skills/

Declare the `agent-harness` plugin under `enabledPlugins`, or version project
review and documentation skills under `.claude/skills/` (D3.3 → ✅).

#### [NEW] .devcontainer/devcontainer.json

Rust devcontainer that installs the toolchain from `rust-toolchain.toml`,
`just` and `pre-commit` (E2 → ✅).

#### [MODIFY] .github/workflows (docs)

Publish `cargo doc --no-deps --workspace` to GitHub Pages (D5, D5.2 → ✅;
counts +0.5 at L2).

#### Re-evaluate

Run the agent-readiness evaluation again to determine the actual level reached.
Report: **Actual level after Phase 1: L? (??%)**

### Phase 2 — Strengthen Level 3

> **Projected level after Phase 2: L3 (L2 set 8 / 10 = 80%)**

#### [MODIFY] sheetlint, sheetstats, sheetcli

Add `tracing` with a `tracing-subscriber` that emits structured output behind
a verbosity flag (O1 → ✅). Code change; needs its own plan.

> [!NOTE]
> E1 (env template) and O2 (error tracking) apply only through the `web-app`
> type (the static WASM demo). The project has no environment variables, so a
> `.env.example` would be empty; decide whether they are worth implementing or
> whether `web-app` should be dropped from the project types.

#### Re-evaluate

Run the agent-readiness evaluation again to determine the actual level reached.
Report: **Actual level after Phase 2: L? (??%)**

### Phase 3 — Path to Level 4 (L3 criteria)

> **Projected level after Phase 3: L4 (L3 set 9 / 10 = 90%)**

#### [MODIFY] .github/workflows/ci.yml

- Pass `--locked` to the cargo commands (B4 → ✅).
- Add a coverage job with `cargo-llvm-cov` and a minimum line threshold
  (T5, T6 → ✅).
- Add a `cargo audit` job (X2 → ✅).

#### [GitHub settings] Ruleset `main`

Release `develop` into `main` through a pull request and require the CI jobs
on `main` too (X5 → ✅; counts at L4).

S7 stays ❌: Rust lints whole crates (see the Pillar 1 note).

#### Re-evaluate

Run the agent-readiness evaluation again to determine the actual level reached.
Report: **Actual level after Phase 3: L? (??%)**

---

## Pros / Cons

| Pros | Cons |
| --- | --- |
| CI enforces fmt, clippy, tests, the WASM build and the no-default-features build on every PR, required to merge into `develop` | `main` does not require the CI checks |
| Dependabot covers Cargo and GitHub Actions with grouping and cooldown | Agent skills no longer versioned in the repo |
| `CLAUDE.md` references shared docs instead of duplicating them | API docs not published |
| Documentation suite complete, indexed and accurate | |
| TDD with positive and negative cases | No coverage in CI, no logging, no devcontainer |

## Verification Plan

### Automated Tests

```zsh
cargo fmt --all -- --check
cargo clippy --workspace --all-targets -- -D warnings
cargo test --workspace
cargo check -p sheetrs-wasm --target wasm32-unknown-unknown
cargo check -p sheetrs --no-default-features
pre-commit run --all-files
```

### Manual Verification

- A PR shows the CI check runs (GitHub UI → PR → Checks).
- A failing check blocks merging into `develop` (ruleset `develop`).
- Dependabot opens a grouped `github-actions` PR on the monthly schedule.
- Every file referenced in `CLAUDE.md` exists.
