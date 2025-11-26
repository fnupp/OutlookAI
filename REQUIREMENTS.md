# Requirements

This document tracks requirements, feature requests, and planned enhancements for OutlookAI.

## Table of Contents
- [Current Requirements](#current-requirements)
- [Completed Requirements](#completed-requirements)
- [Future Considerations](#future-considerations)
- [Non-Functional Requirements](#non-functional-requirements)

## Current Requirements

### High Priority

### [REQ-001] Email Monitoring and Auto-Categorization System

**Priority**: High
**Status**: Planned
**Requested By**: User
**Date**: 2025-11-26

**Description**:
Implement a system that monitors new incoming emails in configured mailboxes and automatically assigns Outlook categories to them using LLM analysis. The system should support multiple categories with configurable LLM prompts and optional automatic draft reply generation per category.

**Acceptance Criteria**:
- [ ] System monitors specified mailboxes for new incoming emails
- [ ] New emails are analyzed by LLM to determine appropriate category
- [ ] Outlook categories are automatically assigned based on LLM response
- [ ] Users can configure which mailboxes to monitor
- [ ] Users can enable/disable the monitoring feature
- [ ] System handles errors gracefully (LLM API failures, network issues)
- [ ] Monitoring runs in background without blocking Outlook UI
- [ ] System provides logging/status information for debugging

**Technical Considerations**:
- Use Outlook `NewMailEx` event or `ItemAdd` event on monitored folders
- Need to handle COM interop for folder monitoring
- Async processing to avoid blocking UI thread
- Consider rate limiting for LLM API calls
- Handle multiple simultaneous new emails efficiently
- Need to persist monitoring state across Outlook restarts

**Related Issues/PRs**:

---

### [REQ-002] Configurable Category Management

**Priority**: High
**Status**: Planned
**Requested By**: User
**Date**: 2025-11-26

**Description**:
Create a configurable list of categories that can be assigned to emails. Each category should support custom naming, LLM prompt configuration for classification, and optional auto-reply settings.

**Acceptance Criteria**:
- [ ] Users can add/edit/delete categories in settings UI
- [ ] Each category has a name (maps to Outlook category name)
- [ ] Each category has a classification prompt for LLM analysis
- [ ] Each category has an optional "generate reply draft" toggle
- [ ] Each category with auto-reply enabled has a configurable reply prompt
- [ ] Category list is persisted in UserData settings
- [ ] UI validates category names (no duplicates, not empty)
- [ ] Changes to categories take effect without restarting Outlook

**Technical Considerations**:
- Extend `UserData.cs` with new properties:
  - `List<EmailCategory> EmailCategories`
  - `bool EmailMonitoringEnabled`
  - `List<string> MonitoredMailboxes`
- Create new `EmailCategory` class with properties:
  - `string CategoryName`
  - `string ClassificationPrompt`
  - `bool GenerateReplyDraft`
  - `string ReplyPrompt`
- Update JSON serialization to handle new nested objects
- Update PromptBox.cs to include category management UI
- Ensure Outlook categories exist before assignment (create if missing)

**Related Issues/PRs**:

---

### [REQ-003] LLM-Based Email Classification

**Priority**: High
**Status**: Planned
**Requested By**: User
**Date**: 2025-11-26

**Description**:
Use LLM to analyze incoming email content and determine which configured category it should be assigned to. The LLM should receive the email subject, body, and sender information along with the list of available categories and their classification prompts.

**Acceptance Criteria**:
- [ ] System extracts relevant email metadata (subject, body, sender, date)
- [ ] System constructs classification prompt with email content and category options
- [ ] LLM response is parsed to determine assigned category
- [ ] Classification prompt is optimized for accurate categorization
- [ ] System handles ambiguous cases (email matches multiple categories)
- [ ] System handles cases where no category matches
- [ ] Classification respects user's LLM provider choice (OpenAI/Ollama)
- [ ] Classification errors are logged and don't crash the add-in

**Technical Considerations**:
- Create master classification prompt template that includes:
  - Email metadata
  - List of categories with their classification prompts
  - Instructions for LLM to return category name
- Use existing `GetLLMResponse` methods from ThisAddIn.cs
- Parse LLM response to extract category name
- Consider using structured output (JSON) for reliable parsing
- Handle HTML email bodies (strip or preserve formatting)
- Limit email body length to avoid token limits
- Consider caching classifications to avoid duplicate API calls

**Related Issues/PRs**:

---

### [REQ-004] Automatic Draft Reply Generation

**Priority**: High
**Status**: Planned
**Requested By**: User
**Date**: 2025-11-26

**Description**:
For categories configured with auto-reply enabled, automatically generate a draft reply email using the category's reply prompt. The draft should be created in the user's Drafts folder and properly linked to the original email as a reply.

**Acceptance Criteria**:
- [ ] Draft replies are only generated for categories with auto-reply enabled
- [ ] Draft is created as a proper reply (RE: subject, correct recipients)
- [ ] LLM generates reply content based on category's reply prompt and original email
- [ ] Draft is saved to Drafts folder
- [ ] Draft maintains conversation threading with original email
- [ ] User can manually review and edit draft before sending
- [ ] System handles multiple simultaneous draft generations
- [ ] Reply generation failures don't prevent category assignment

**Technical Considerations**:
- Use `MailItem.Reply()` or `MailItem.ReplyAll()` to create draft
- Set `MailItem.Body` or `HTMLBody` with LLM-generated content
- Don't call `Send()` - leave in drafts for user review
- Reply prompt should include original email context
- Consider using `§§Input§§` placeholder pattern for consistency
- Handle edge cases:
  - External senders (reply vs reply all)
  - Distribution lists
  - Multiple recipients
- Properly release COM objects for draft items

**Related Issues/PRs**:

---

### [REQ-005] Settings UI for Email Monitoring

**Priority**: High
**Status**: Planned
**Requested By**: User
**Date**: 2025-11-26

**Description**:
Extend the PromptBox settings dialog to include configuration options for email monitoring, category management, and mailbox selection.

**Acceptance Criteria**:
- [ ] New tab or section in PromptBox for email monitoring settings
- [ ] Enable/disable toggle for email monitoring feature
- [ ] Mailbox selection UI (list of available mailboxes with checkboxes)
- [ ] Category management UI:
  - List/grid showing all configured categories
  - Add/Edit/Delete buttons for categories
  - Category detail form with all properties
- [ ] Input validation and error messages
- [ ] Preview/test functionality to test classification
- [ ] Save/Cancel buttons persist changes to UserData

**Technical Considerations**:
- Add new TabPage to existing PromptBox TabControl
- Use DataGridView or ListView for category list
- Create separate form or panel for category editing
- Enumerate Outlook folders/mailboxes using Outlook Interop
- Handle different folder types (IMAP, Exchange, POP3)
- Use data binding to UserData where possible
- Maintain existing settings UI patterns for consistency

**Related Issues/PRs**:

---

### Medium Priority


### Low Priority


## Completed Requirements

Document completed requirements here with completion date and version.

| Requirement | Description | Completed Date | Version |
|-------------|-------------|----------------|---------|
| | | | |

## Future Considerations

Ideas and potential features that need further discussion or planning.

-

## Non-Functional Requirements

### Performance
- Email classification should not block Outlook UI (async processing required)
- LLM API calls should be rate-limited to avoid overwhelming the service
- System should handle high volumes of incoming emails efficiently
- Memory usage should remain reasonable with large mailboxes
- Category assignment should complete within 10 seconds per email under normal conditions

### Security
- API keys must be stored securely
- Proxy passwords encrypted using Windows DPAPI
- Email content sent to LLM should respect user privacy expectations
- Consider option to exclude sensitive emails from auto-processing
- Ensure external LLM providers are accessed over secure connections (HTTPS)

### Compatibility
- Windows OS with .NET Framework 4.8
- Microsoft Outlook 365 (classic)
- Office Interop v15.0
- Must work with various mailbox types (Exchange, IMAP, POP3)
- Support multiple Outlook profiles/accounts

### Usability
- Settings UI should be intuitive and easy to navigate
- Category management should follow familiar UI patterns
- Errors should be communicated clearly to users
- Users should be able to easily enable/disable monitoring without losing configuration
- Provide clear feedback when emails are being processed

### Maintainability
- Code should follow existing project patterns and conventions
- Email monitoring logic should be modular and testable
- Logging should be comprehensive for troubleshooting
- Settings schema changes should handle migration from older versions

---

## Requirement Template

When adding new requirements, use this template:

```markdown
### [REQ-XXX] Requirement Title

**Priority**: High/Medium/Low
**Status**: Planned/In Progress/Completed/Deferred
**Requested By**: [Name/Role]
**Date**: YYYY-MM-DD

**Description**:
Clear description of what is needed.

**Acceptance Criteria**:
- [ ] Criterion 1
- [ ] Criterion 2
- [ ] Criterion 3

**Technical Considerations**:
Any technical notes, constraints, or dependencies.

**Related Issues/PRs**: #issue-number, #pr-number
```
