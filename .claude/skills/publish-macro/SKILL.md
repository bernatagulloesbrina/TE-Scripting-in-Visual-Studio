---
name: publish-macro
description: Publish a Tabular Editor macro (a void method in "TE Scripts/TE Scripts.cs") to the local TabularEditor-Scripts distribution repo - run CopyMacro for the method and save the code, with its // MACRO header, as a new file in the right folder. Use when the user says "publish <METHOD NAME>", "publish the macro <X>", or asks to copy a macro into the publication / distribution repo.
---

# Publish a macro

The publication repo is a local clone of github.com/bernatagulloesbrina/TabularEditor-Scripts. The te-macro-skills importer (te-macros plugin) installs macros from its files, reading the `// MACRO:` header of each one, so every published file must carry that header.

## Step 0 - Locate the publication repo

Use, in this order:

1. `$env:TE_PUBLISH_REPO`, if set.
2. `..\TabularEditor-Scripts` next to this repo (the solution folder), if it exists.
3. Otherwise explain that publish-macro needs a local clone of TabularEditor-Scripts (point them to the "One-time setup" section of the repo's README.md), ask for the path, and offer to save it in `.claude\settings.local.json` (per machine, not committed) so they aren't asked again:

   ```json
   { "env": { "TE_PUBLISH_REPO": "D:\\repos\\TabularEditor-Scripts" } }
   ```

   Variables set there apply from the next session on.

Check that the folder has `Advanced\Skill-Ready Macros` and `plugins\te-macros` before using it. Below, `<publication repo>` is that folder.

## Step 1 - Find the method

Find `<METHOD NAME>` in `TE Scripts/TE Scripts.cs`. Macros are the methods that are not `public static` (CopyMacro only lists those). If there is no exact match, show the closest names and ask.

Read the method's header comments (after its `//using` / `//#r` lines).

## Step 2 - Pick the folder

**Skill-ready**: the method has a complete `// SKILL:` header, meaning `// SKILL: name:`, `// SKILL: description:` and the body with `## Inputs (environment variables)` (see the make-skill-ready skill), plus a `#if TECLI` block. Publish it to `Advanced\Skill-Ready Macros` as `<MACRO Name>.cs` without asking.

**Not skill-ready** (no `// SKILL:` lines): list the folders of the publication repo (`Basic`, `Intermediate`, `Advanced` and its subfolders, skipping `plugins` and `Tabular Editor 2`, which is for TE2 scripts) and ask the user where to publish it with AskUserQuestion. Skill-Ready Macros isn't an option for these. The file is `<MACRO Name>.csx`, matching the other files in those folders. Also offer to make it skill-ready instead (make-skill-ready skill).

**Partial `// SKILL:` header** (some lines missing): don't publish. Say what is missing and offer to complete it with make-skill-ready.

The file name is always the friendly name from `// MACRO: Name:` (Step 3), never the method name. Replace characters Windows doesn't allow in file names (`\ / : * ? " < > |`) with `-`.

If a file for this macro is already in the repo (same file name, or the same `// MACRO: Name:` / `// SKILL: name:` in any file), this is an update. Write the new code to `<MACRO Name>.<ext>` in the chosen folder. If the existing file has another name or folder, move it there with `git mv` first so its history is kept, then overwrite it, and tell the user.

## Step 3 - Make sure the MACRO header is there

The method needs these three lines, right after its `//using` / `//#r` lines and before the `// SKILL:` lines and changelog comments:

```csharp
// MACRO: Name: <friendly macro name shown in TE, title case>
// MACRO: ValidContexts: <contexts as in MacroActions.json: Model, Table, Column, Measure, Function, Calendar...; several separated by ", ">
// MACRO: Tooltip: <one sentence: what the macro does>
```

If any of them are missing, add them **to the method in TE Scripts.cs** (not only to the published file), so later publishes keep them. Derive ValidContexts from the `Selected.*` the method uses (`Model` if none) and the Tooltip from what it does. Show the user the values you chose with the folder question, or before publishing if you aren't asking about a folder. The lines must not contain `NOCOPY` (CopyMacro drops those lines).

## Step 4 - Run CopyMacro

```powershell
& "${CLAUDE_SKILL_DIR}\scripts\Copy-Macro.ps1" -Method <METHOD NAME> -OutFile "<publication repo>\<folder>\<file>"
```

It builds the `TE Scripts` project and calls `TE_Scripts.CopyMacroFromVSFile` from the built dll, passing the method name and the output file. The code is the same CopyMacro would put on the clipboard in TE: directives merged, GeneralFunctions / ReportFunctions / Report / DaxUserDefinedFunction classes appended when the method uses them, and the helpers the CLI branch doesn't reach wrapped in `#if !TECLI`. Publish that output as is; don't hand-edit the published file.

## Step 5 - Check

- The file starts with the directives, then the `// MACRO:` lines (and the `// SKILL:` lines if skill-ready). The directives are plain lines, never inside `#if` / `#endif`: TE3 moves them out of the macro and the leftover `#endif` stops every macro in MacroActions.json from loading.
- Preview the import: `& "<publication repo>\plugins\te-macros\skills\te-macro-skills\scripts\Import-Macros.ps1" "<published file>" -WhatIf`. Status must be `created` or `updated` with no Problems. `-WhatIf` writes nothing.
- Skill-ready macros: `te script --file "<published file>" --validate --non-interactive` must print `Compilation OK`. If `te` isn't on PATH, use `$env:LOCALAPPDATA\Programs\TabularEditor3\te.exe`.
- For an update, show `git -C <publication repo> diff --stat`.

## Step 6 - Report

Tell the user the file path, the folder, whether it was new or an update, any `// MACRO:` lines you added to TE Scripts.cs, and the check results. Don't commit or push either repo unless the user asks. When they do, commit the publication repo on a feature branch, not `main`.
