# AI Outlook Add-in

Welcome to the AI Outlook Add-in, a powerful tool that enhances your email workflow with LLM capabilities. It integrates directly into classic Outlook 365 as a VSTO add-in, supporting both OpenAI API and Ollama for AI-powered features.

If you appreciate my work, please consider giving it a star! 🤩

## Features

- **AI-Powered Email Responses**: Generate intelligent replies to emails using customizable prompts
- **Email Summarization**: Summarize single or multiple emails quickly
- **Text Composition Assistant**: Improve and rewrite email text while composing
- **Email Monitoring & Auto-Categorization**: Automatically categorize incoming emails using LLM analysis
- **Auto-Reply Draft Generation**: Generate draft replies automatically for specific categories
- **Flexible LLM Support**: Works with OpenAI API or local Ollama models
- **Proxy Support**: Built-in proxy configuration with encrypted password storage

# Setup
1. Download the release
2. Quit Outlook
3. Install by starting OutlookAI.vsto
4. Restart Outllok

# Update
1. Download the release
2. Quit Outlook
3. Install by starting OutlookAI.vsto.
   Due to the way vtso installations works please make sure that you start the installation from the same directory as the installation. If the installation fails - deinstall via the Windows Settings and reinstall. The application seetings will remain.
4. Restart Outllok

# Deinstall
1. Quit Outlook
2. Deinstall via the Windows Settings

# Configuration

## Basic Setup
In the settings, configure your OpenAI or Ollama information:
![grafik](https://github.com/user-attachments/assets/72c77bd0-9058-4931-9786-566be29cce56)

If needed, configure your proxy settings:
![grafik](https://github.com/user-attachments/assets/3b7f9d1b-a98f-4109-a660-f16a2fa30aa5)

## Email Monitoring Setup
1. Open OutlookAI settings and navigate to the "Email Monitoring" tab
2. Check "Enable Email Monitoring"
3. Select the mailboxes you want to monitor
4. Add categories with classification prompts
5. Optionally enable auto-reply draft generation for specific categories
6. Save settings and restart Outlook if needed

# Documentation

For detailed information about the project:
- **Development Guide**: See [CLAUDE.md](CLAUDE.md) for build instructions and architecture overview
- **Requirements**: See [REQUIREMENTS.md](REQUIREMENTS.md) for complete feature specifications
- **Implementation Notes**: See [IMPLEMENTATION_NOTES.md](IMPLEMENTATION_NOTES.md) for technical details about email monitoring

# Tech Stack

- .NET Framework 4.8
- VSTO (Visual Studio Tools for Office)
- Microsoft Office Interop (Outlook, Word) v15.0
- Newtonsoft.Json 13.0.3
- OpenAI API / Ollama support

# ToDo

1. Translations (localization for additional languages)






