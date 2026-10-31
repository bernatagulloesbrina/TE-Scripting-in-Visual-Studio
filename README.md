# TE-Scripting-in-Visual-Studio
B.Agullo 2022-12-22
www.esbrina-ba.com/blog
Please see 
https://www.esbrina-ba.com/c-scripting-nirvana-effortlessly-use-visual-studio-as-your-coding-environment/
for details

## Publishing macros with Claude Code

The repo includes two Claude Code skills in `.claude/skills`:

- **make-skill-ready**: adds the `// MACRO:` and `// SKILL:` header to a script and makes it run in the Tabular Editor CLI (`te`).
- **publish-macro**: when you say `publish <METHOD NAME>`, it runs CopyMacro for that method and saves the code to a local clone of [TabularEditor-Scripts](https://github.com/bernatagulloesbrina/TabularEditor-Scripts), the repo the macros are distributed from.

### One-time setup: tell Claude where your TabularEditor-Scripts clone is

publish-macro writes into your clone of TabularEditor-Scripts, so it needs to know where that clone is on your machine. It looks in this order:

1. The `TE_PUBLISH_REPO` environment variable.
2. A `TabularEditor-Scripts` folder next to this repo, for example:

   ```
   C:\repos\TE-Scripting-in-Visual-Studio   <- this repo
   C:\repos\TabularEditor-Scripts           <- found automatically
   ```

3. If neither is found, Claude asks you for the path and offers to save it for you.

If your clone isn't next to this repo, set `TE_PUBLISH_REPO` once. Create `.claude/settings.local.json` in this repo with:

```json
{
  "env": {
    "TE_PUBLISH_REPO": "D:\\path\\to\\TabularEditor-Scripts"
  }
}
```

- Use double backslashes (`\\`) or forward slashes in the path.
- If the file already exists, add the `env` block to it rather than replacing the file.
- Start a new Claude Code session afterwards. Settings are read at startup.
- The file is in `.gitignore`, so your path stays on your machine. To use the same path in every project, put the `env` block in `%USERPROFILE%\.claude\settings.json` instead.

### Other requirements

- The .NET SDK (`dotnet`), to build the `TE Scripts` project. publish-macro builds it before every publish.
- The Tabular Editor CLI (`te`), on PATH or in `%LOCALAPPDATA%\Programs\TabularEditor3\`. It's used to check that skill-ready macros compile.
