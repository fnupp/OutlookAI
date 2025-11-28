# Requirements

This document tracks requirements, feature requests, and planned enhancements for OutlookAI.

## Table of Contents
- [Current Requirements](#current-requirements)
- [Completed Requirements](#completed-requirements)
- [Future Considerations](#future-considerations)
- [Non-Functional Requirements](#non-functional-requirements)

## Current Requirements

### High Priority

_No open high-priority requirements at this time._

## Completed Requirements

### [REQ-001] Email Monitoring and Auto-Categorization System

**Priority**: High
**Status**: ✅ Completed
**Requested By**: User
**Date**: 2025-11-26
**Completed**: 2025-11-27

**Description**:
Implement a system that monitors new incoming emails in configured mailboxes and automatically assigns Outlook categories to them using LLM analysis. The system should support multiple categories with configurable LLM prompts and optional automatic draft reply generation per category.

**Acceptance Criteria**:
- [x] System monitors specified mailboxes for new incoming emails
- [x] New emails are analyzed by LLM to determine appropriate category
- [x] Outlook categories are automatically assigned based on LLM response
- [x] Users can configure which mailboxes to monitor
- [x] Users can enable/disable the monitoring feature
- [x] System handles errors gracefully (LLM API failures, network issues)
- [x] Monitoring runs in background without blocking Outlook UI
- [x] System provides logging/status information for debugging

**Technical Considerations**:
- Use Outlook `NewMailEx` event or `ItemAdd` event on monitored folders
- Need to handle COM interop for folder monitoring
- Async processing to avoid blocking UI thread
- Consider rate limiting for LLM API calls
- Handle multiple simultaneous new emails efficiently
- Need to persist monitoring state across Outlook restarts

---

### [REQ-002] Configurable Category Management

**Priority**: High
**Status**: ✅ Completed
**Requested By**: User
**Date**: 2025-11-26
**Completed**: 2025-11-27

**Description**:
Create a configurable list of categories that can be assigned to emails. Each category should support custom naming, LLM prompt configuration for classification, and optional auto-reply settings.

**Acceptance Criteria**:
- [x] Users can add/edit/delete categories in settings UI
- [x] Each category has a name (maps to Outlook category name)
- [x] Each category has a classification prompt for LLM analysis
- [x] Each category has an optional "generate reply draft" toggle
- [x] Each category with auto-reply enabled has a configurable reply prompt
- [x] Category list is persisted in UserData settings
- [x] UI validates category names (no duplicates, not empty)
- [x] Changes to categories take effect without restarting Outlook

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

---

### [REQ-003] LLM-Based Email Classification

**Priority**: High
**Status**: ✅ Completed
**Requested By**: User
**Date**: 2025-11-26
**Completed**: 2025-11-27

**Description**:
Use LLM to analyze incoming email content and determine which configured category it should be assigned to. The LLM should receive the email subject, body, and sender information along with the list of available categories and their classification prompts.

**Acceptance Criteria**:
- [x] System extracts relevant email metadata (subject, body, sender, date)
- [x] System constructs classification prompt with email content and category options
- [x] LLM response is parsed to determine assigned category
- [x] Classification prompt is optimized for accurate categorization
- [x] System handles ambiguous cases (email matches multiple categories)
- [x] System handles cases where no category matches
- [x] Classification respects user's LLM provider choice (OpenAI/Ollama)
- [x] Classification errors are logged and don't crash the add-in

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

---

### [REQ-004] Automatic Draft Reply Generation

**Priority**: High
**Status**: ✅ Completed
**Requested By**: User
**Date**: 2025-11-26
**Completed**: 2025-11-27

**Description**:
For categories configured with auto-reply enabled, automatically generate a draft reply email using the category's reply prompt. The draft should be created in the user's Drafts folder and properly linked to the original email as a reply.

**Acceptance Criteria**:
- [x] Draft replies are only generated for categories with auto-reply enabled
- [x] Draft is created as a proper reply (RE: subject, correct recipients)
- [x] LLM generates reply content based on category's reply prompt and original email
- [x] Draft is saved to Drafts folder
- [x] Draft maintains conversation threading with original email
- [x] User can manually review and edit draft before sending
- [x] System handles multiple simultaneous draft generations
- [x] Reply generation failures don't prevent category assignment

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

---

### [REQ-005] Settings UI for Email Monitoring

**Priority**: High
**Status**: ✅ Completed
**Requested By**: User
**Date**: 2025-11-26
**Completed**: 2025-11-27

**Description**:
Extend the PromptBox settings dialog to include configuration options for email monitoring, category management, and mailbox selection.

**Acceptance Criteria**:
- [x] New tab or section in PromptBox for email monitoring settings
- [x] Enable/disable toggle for email monitoring feature
- [x] Mailbox selection UI (list of available mailboxes with checkboxes)
- [x] Category management UI:
  - List/grid showing all configured categories
  - Add/Edit/Delete buttons for categories
  - Category detail form with all properties
- [x] Input validation and error messages
- [x] Preview/test functionality to test classification
- [x] Save/Cancel buttons persist changes to UserData

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

_No open medium-priority requirements at this time._

### Low Priority

_No open low-priority requirements at this time._

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
