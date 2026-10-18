# AGENTS.md

Guidance for AI coding agents (Claude Code, Copilot, etc.) working in this repository.

## What this is

An **OutSystems Developer Cloud (ODC) External Logic** library written in C# (.NET 10).
ODC loads the compiled DLL and exposes each method of the `[OSInterface]` interface as a
server action that ODC apps can call. It currently has a single sample action, `Greet`.

## Layout

```
src/OdcBatchExcel/            The library (uploaded to ODC)
  IOdcBatchExcel.cs           [OSInterface] contract: one [OSAction] per server action
  OdcBatchExcel.cs            Implementation
tests/OdcBatchExcel.Tests/    xUnit v3 tests (Microsoft.Testing.Platform runner)
scripts/package.sh            Builds the ODC-ready zip into artifacts/
Directory.Build.props         Shared build + analyzer settings (all projects)
Directory.Packages.props      Central NuGet versions (set versions here, never in .csproj)
.editorconfig                 Formatting and analyzer severities
global.json                   Pins the .NET SDK and opts into the MTP test runner
.husky/                       Pre-commit hook (Husky.Net)
.github/                      CI workflow + Dependabot
```

## Commands

Run everything from the repository root.

| Task | Command |
|---|---|
| Restore tools + packages | `dotnet tool restore && dotnet restore` |
| Build (lint + static analysis) | `dotnet build -c Release` |
| Format code | `dotnet format` |
| Check formatting (CI) | `dotnet format --verify-no-changes` |
| Run tests | `dotnet test -c Release` |
| Tests + coverage (Cobertura) | `dotnet test -c Release --coverage --coverage-output-format cobertura --coverage-output coverage.cobertura.xml` |
| Build the ODC zip | `scripts/package.sh` → `artifacts/OdcBatchExcel.zip` |

Before you call a change done, run: `dotnet format`, `dotnet build -c Release` and `dotnet test -c Release`.

## Quality gates

- **Warnings are errors.** `TreatWarningsAsErrors`, `AnalysisLevel=latest-recommended` and
  `EnforceCodeStyleInBuild` are on. The analyzers are the built-in .NET (CA/IDE) ones,
  **StyleCop.Analyzers** and **SonarAnalyzer.CSharp**. Fix the cause. Only suppress a rule in
  `.editorconfig`, with a comment explaining why, never with blanket `#pragma`s.
- Public members in `src/` need XML doc comments (they also become the descriptions in ODC).
- StyleCop wants `this.` before instance members, and `using` directives outside the namespace.
- **Pre-commit hook**: `dotnet format --verify-no-changes` runs on staged `.cs` files. It is
  installed automatically on the first local build (`dotnet husky install`). Set `HUSKY=0` to skip it.
- **CI** (`.github/workflows/ci.yml`) runs the format check, build, tests with coverage and
  packaging on every push to `main` and on every PR, then uploads the zip and the coverage report as artifacts.

## ODC External Logic rules

- **Never rename `[OSInterface(Name = "OdcBatchExcel")]`.** The name is the library's identity
  in the tenant: a new name creates a separate library, and apps that use the old one break.
  A test (`Interface_KeepsStableOdcLibraryName`) enforces this.
- Add actions as methods on `IOdcBatchExcel` with `[OSAction]`, and add `[OSParameter]` to
  parameters. Implement them in `OdcBatchExcel`. Renaming or removing an action, or changing
  its signature, breaks the apps that use it.
- Use ODC-compatible types: `string`, `int`, `long`, `decimal`, `double`, `bool`,
  `DateTime`, `byte[]` (Binary Data), structs marked `[OSStructure]` (fields with
  `[OSStructureField]`), and `List<T>` of those. Use `[OSIgnore]` to hide a public member.
- The library is **stateless**: no static mutable state, and don't rely on anything left over
  from a previous call. Throwing an exception surfaces as an error in the calling ODC app.
- The zip must contain `OdcBatchExcel.dll` and `OdcBatchExcel.deps.json` **at the root**
  (no nested folder) and be ≤ 50 MB. `scripts/package.sh` checks both.

## Deploying to ODC (OutSystems MCP tools)

A library goes through four steps:

1. **Upload** (MCP): `scripts/package.sh` → base64-encode the zip → `extlib_upload`
   (`zip_b64`) → poll `extlib_status` until `ReadyForReview`. On failure, read
   `extlib_logs`. Check the detected actions with `extlib_contents`.
2. **Publish** (MCP): `extlib_publish` → poll `extlib_status` until `Published`.
   Publishing is **not** the official release.
3. **Try library in an app** (ODC Portal, done by the maintainer): add the published revision
   to an app and test its actions by hand.
4. **Release Library** (ODC Portal, done by the maintainer): makes the revision the official
   version. No MCP tool can do steps 3 or 4. Stop after step 2 and let the maintainer take over.

**Uploading, publishing and `extlib_delete` change tenant state. Always get the user's
explicit confirmation first.** The status, logs, contents and listing tools are read-only.

## Git

Commit when asked, but do **not** push. The maintainer runs `git push` themselves.
