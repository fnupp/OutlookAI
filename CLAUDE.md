# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OutlookAI is a VSTO (Visual Studio Tools for Office) add-in for Microsoft Outlook 365 (classic) that integrates LLM capabilities directly into the email workflow. It supports both OpenAI API and Ollama for AI-powered email responses, summarization, and text composition.

## Build and Development Commands

### Building the Project
```bash
# Build in Debug mode
msbuild OutlookAI.sln /p:Configuration=Debug

# Build in Release mode
msbuild OutlookAI.sln /p:Configuration=Release
```

### Installing the Add-in for Testing
1. Close Outlook completely
2. Build the project in Debug or Release mode
3. Navigate to `OutlookAI\bin\Debug\` (or `Release`)
4. Run `OutlookAI.vsto` to install
5. Restart Outlook

### Debugging
- The project is configured to launch Outlook when debugging (F5 in Visual Studio)
- Debug path is set to Office 16.0 Outlook executable
- Breakpoints can be set in any C# code file

## Git Workflow

### Branching Strategy
**IMPORTANT**: When generating or modifying code, ALWAYS create a feature branch first. Never commit code changes directly to `master`.

#### Workflow Steps:
1. **Before making any code changes**, create a new branch from `master`:
   ```bash
   git checkout master
   git pull origin master
   git checkout -b <branch-name>
   ```

2. **Make all code changes** on the feature branch

3. **Commit changes** with descriptive commit messages

4. **Push to remote** (if ready for review):
   ```bash
   git push -u origin <branch-name>
   ```

5. **Create a Pull Request** for code review before merging to `master`

#### Branch Naming Conventions:
Use descriptive names that clearly indicate the purpose of the branch:

- `feature/<description>` - For new features
  - Example: `feature/add-usage-tracking`
  - Example: `feature/multi-language-support`

- `fix/<description>` - For bug fixes
  - Example: `fix/email-body-encoding`
  - Example: `fix/ribbon-button-visibility`

- `enhancement/<description>` - For improvements to existing features
  - Example: `enhancement/improve-error-handling`
  - Example: `enhancement/optimize-llm-calls`

- `refactor/<description>` - For code refactoring
  - Example: `refactor/ribbon-ui-components`
  - Example: `refactor/settings-management`

#### Important Notes:
- Keep `master` branch clean and production-ready
- One feature/fix per branch (avoid mixing unrelated changes)
- Delete branches after they are merged
- Always pull latest `master` before creating a new branch

## Architecture

### Core Components

**ThisAddIn.cs** - VSTO Add-in entry point
- Initializes the add-in on Outlook startup
- Manages user settings stored in `%AppData%\OutlookAI\OutlookAI.json`
- Provides centralized LLM communication methods (`GetLLMResponse`, `GetChatGPTResponse`, `GetChatOllamaResponse`)
- Handles HTTP client creation with optional proxy support

**UserData.cs** - Settings data model
- Stores all user-configurable settings (prompts, API keys, model names, proxy settings)
- Implements encrypted password storage using Windows DPAPI (`ProtectedData.Protect`)
- Serialized to/from JSON for persistence

### Ribbon UI Components

The add-in uses Outlook Ribbon UI with two contexts:

**OutlookAIRibbon.cs** - Mail reading/viewing context
- Appears when viewing received emails
- Provides 4 configurable prompt buttons for AI-powered reply generation
- Implements email summarization (single or multiple emails)
- Includes calendar export/import functionality (JSON format)
- Methods: `Reply()`, `GetMail()`, `GetMails()`, `Summarize()`, `SummarizeMultiple()`

**ComposeRibbon.cs** - Mail composition context
- Appears when composing new emails
- Provides 3 configurable prompt buttons for text improvement/rewriting
- Works with selected text in the Word editor
- Methods: `GetSelectedText()`, `UpdateMail()`, button click handlers

### Dialog Forms

**PromptBox.cs** - Settings dialog
- Edits all user configuration (prompts, titles, API settings, proxy settings)
- Uses data binding to `UserData` via `userDataBindingSource`
- Fetches available Ollama models from Ollama API
- Saves settings back to JSON file on OK

**InputBox.cs** - Simple text input dialog
- Prompts user for additional input when using reply/compose prompts
- Used to collect user instructions that get inserted into prompts via `§§Input§§` placeholder

## Key Technical Details

### LLM Integration
- **Ollama**: POST to `{OllamaUrl}/api/generate` with `{model, prompt, stream: false}`
- **OpenAI**: POST to OpenAI-compatible API with bearer token authentication
- Both use async/await pattern with `HttpClient`
- Prompt template: Prompts can include `§§Input§§` placeholder for user input

### Settings Storage
- Location: `%AppData%\OutlookAI\OutlookAI.json`
- Initialized on first run with default localized prompts from `Resources.resx`
- Proxy passwords encrypted per-user using Windows DPAPI

### Localization
- German (de) and English resource files
- Target culture set to "de" in project file
- All UI labels pulled from `Resources.resx` and `Resources.de.resx`

### COM Interop
- Uses Office 15.0 Interop assemblies (Outlook, Word)
- Properly releases COM objects using `Marshal.ReleaseComObject()`
- Forces garbage collection after COM-intensive operations

### Email Context Handling
- In reading context: Gets mail from `Inspector.CurrentItem` or `ActiveExplorer().Selection`
- In compose context: Gets current `MailItem` from `Inspector` context
- Uses Word editor for text selection in compose mode

## Dependencies
- .NET Framework 4.8
- Microsoft Office Interop (Outlook, Word) v15.0
- Newtonsoft.Json 13.0.3
- Visual Studio Tools for Office Runtime 4.0

## Manifest Signing
- Project uses ClickOnce signing with temporary PFX certificates
- Current key: `OutlookAI_3_TemporaryKey.pfx`
- Thumbprint: `5DAC2D935FB45E072E3FF373BA97419E8D6CAD3A`

## Publication Settings
- Target culture: German (de)
- Version: 1.0.0.31 (auto-incrementing)
- Update interval: 7 days
- Bootstrap: .NET Framework 4.8 and VSTO Runtime required
