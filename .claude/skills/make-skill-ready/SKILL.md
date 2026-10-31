---
name: make-skill-ready
description: Make a Tabular Editor script (a void method in "TE Scripts/TE Scripts.cs") skill-ready - add the // MACRO and // SKILL header and route every user input through environment variables in CLI mode with #if TECLI. Use when the user asks to make a script/method/macro "skill-ready", "agent-ready", "CLI-ready", or to add the skill header to a script.
---

# Make a script skill-ready

A skill-ready script runs unchanged in Tabular Editor (interactive) and, when compiled with `TECLI` (te CLI), takes every input from environment variables instead of the selection or dialogs. Its header carries the macro metadata (for auto-import into MacroActions.json) and the skill text (a generator turns it into a Claude skill per installation).

Reference implementation: the `SUM` method at the top of `TE Scripts/TE Scripts.cs`. Read it first.

## Rules

- Modify only the target method (plus helpers in GeneralFunctions / ReportFunctions when needed). Follow `.github/copilot-instructions.md`.
- `#if TECLI` appears **only** in TE Scripts.cs, never in GeneralFunctions or ReportFunctions.
- In the `TECLI` branch, never touch `Selected`, dialogs (`Interaction.InputBox`, `MessageBox`, `SelectTable`, `Fx.ChooseString`, ...) or anything that prompts.
- The `#else` branch keeps the original interactive code exactly as it was.
- Header lines must not contain `NOCOPY` (CopyMacro drops those lines) and must not start with `//using` or `//#r`.

## Step 1 - Understand the script

Read the whole method and the Fx / Rx helpers it calls (GeneralFunctions/GeneralFunctions.cs, ReportFunctions/ReportFunctions.cs); helpers often prompt internally. Work out:
- what it does, what it creates or changes, and its side effects (hides columns, deletes objects, writes report files...)
- every user input: `Selected.*` and every prompt, including those hidden inside helpers
- for each input: its type, whether it's required, and its default value if there is one (default names usually live in a variable or in the `DefaultResponse` passed to `Fx.GetNameFromUser`)

## Step 2 - Replace inputs with TECLI blocks

Declare each input variable once, fill it in a `#if TECLI` / `#else` / `#endif` block, then use only that variable afterwards (replace every later `Selected.X` use). Group the inputs in one block near the start of the method when the code allows it. When an input depends on an earlier one (for example a column of a selected table), use a block where it is asked.

Env variable names: `<PREFIX>_<INPUT>` in upper case, where the prefix is a short form of the skill name, e.g. `SUM_COLUMNS`, `FP_NAME`, `FP_MEASURES`.

Use the ENV helpers in `Fx` (GeneralFunctions). Each one shows an Error and returns null if the variable is missing or can't be resolved, so follow every call with `if (x == null) return;`.

| Interactive input | TECLI replacement |
|---|---|
| `Selected.Columns` / `SelectColumn(...)` | `Fx.GetColumnsFromEnv(Model, "X")` / `Fx.GetColumnFromEnv(Model, "X")` |
| `Selected.Measures` / `SelectMeasure(...)` / `Fx.GetSelectedMeasure(...)` | `Fx.GetMeasuresFromEnv` / `Fx.GetMeasureFromEnv` |
| `Selected.Tables` / `SelectTable(...)` / `Fx.GetDateTable` / `Fx.GetFactTable` | `Fx.GetTablesFromEnv` / `Fx.GetTableFromEnv` |
| `Selected.Functions` | `Fx.GetFunctionsFromEnv` |
| `Fx.GetNameFromUser(..., DefaultResponse: d)` / `Interaction.InputBox` | `Fx.GetStringFromEnv("X", defaultValue: d)` (parse numbers as the original code does) |
| `Fx.ChooseString(options)` | `Fx.GetStringFromEnv("X")`, then check `options.Contains(value)` and Error if not |
| `Fx.ChooseStringMultiple` / `Fx.Select*Multiple` (return name lists) | `Fx.GetStringsFromEnv("X")` (same name format the original returns, e.g. DaxObjectFullName) |
| `Fx.IsAnswerYes(...)` / Yes-No `MessageBox` | `Fx.GetBoolFromEnv("X", defaultValue: ...)`, null check, then `.Value` |

If an input has no helper, or a helper prompts internally (e.g. `Rx.InitReport` asks for the .pbir path), add an overload or a new `...FromEnv` helper to GeneralFunctions / ReportFunctions that takes the value as a parameter, following the same pattern and with no `#if`. Call it from the TECLI branch.

Pure messages (`Info`, `Output`, `Warning`) can stay. A confirmation that only shows information and has no real choice can be skipped in the TECLI branch.

### Default names

Names for new objects (measures, field parameters, calc groups, tables...) usually have defaults. Keep the default in a variable so both branches use it:

```csharp
string defaultName = "Field Parameter";
#if TECLI
            string fpName = Fx.GetStringFromEnv("FP_NAME", defaultValue: defaultName);
#else
            string fpName = Fx.GetNameFromUser(Prompt: "Field parameter name", Title: "Name", DefaultResponse: defaultName);
#endif
            if (fpName == null) return;
```

The variable is then optional. In the skill text, list it as optional, copy the default value literally, and say when the default is fine (it is in most cases) and when to set it instead: the user asked for a specific name, the default would clash with an existing object, or the script is being run several times and the names must differ.

## Step 3 - Write the header

Put it right after the method's `//using` / `//#r` lines and before the changelog comments, and add a changelog line for the change:

```csharp
// MACRO: Name: <friendly macro name shown in TE>
// MACRO: ValidContexts: <comma-separated contexts>
// MACRO: Tooltip: <one sentence: what the macro does>
// SKILL: name: <kebab-case skill name, stable, never reuse>
// SKILL: description: <what it does + when to use it, with trigger phrases a user would say>
// SKILL:
// SKILL: <short paragraph: what the script does, what it creates/changes, side effects>
// SKILL:
// SKILL: ## Inputs (environment variables)
// SKILL: Set these variables before running the macro. Work out each value from the conversation and the model (inspect the model if needed); if a value is unclear, ask the user instead of guessing.
// SKILL: - <VAR> (required): <format, e.g. 'Table'[Column] refs separated by ';'>. <what it is>
// SKILL: - <VAR> (optional, default "<literal default>"): <what it is>. <when to keep the default and when to set it>
// SKILL:
// SKILL: ## Notes
// SKILL: <preconditions (compatibility level, date table, report open...), side effects, anything the agent should check or tell the user>
```

- **Name**: a friendly name in title case; it doesn't have to match the method name.
- **ValidContexts**: the object types the script acts on in TE, as written in MacroActions.json (`Model`, `Table`, `Column`, `Measure`, `Function`, `Calendar`, ...). Derive them from the `Selected.*` the script uses; use `Model` when it uses no selection. Several values are separated by `, `, e.g. `Measure, Column`.
- **SKILL name**: lowercase letters, digits and hyphens. It links the skill to the macro across installations, so once published never change it.
- **SKILL description**: one line, written for an agent deciding whether to use the skill.

## Step 4 - Check

- Both branches compile in your head: same variable names and types after `#endif`, and no `Selected` or prompt in the TECLI branch.
- Use the IDE diagnostics (the Visual Studio project compiles the `#else` branch; TECLI isn't defined there).
- Every env variable used in the code is in the `## Inputs` list, with matching names, and the reverse.
- te has no WinForms / VisualBasic. The TECLI branch, and every Fx / Rx helper it calls (directly or through other helpers), must not use them. CopyMacro wraps the helpers the TECLI branch doesn't reach in `#if !TECLI`, so dialog helpers used only by the `#else` branch are fine.
- Once the user has copied the macro with CopyMacro and saved it in TE3, the te-macro-skills skill (te-macros plugin) checks the header and compiles the macro in te: generate with `-Only <skill name> -WhatIf`.
- Report to the user: the env variables, their defaults, the helpers you added, and anything you couldn't make non-interactive.
