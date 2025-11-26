using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI
{
    /// <summary>
    /// Service that monitors configured mailboxes for new emails and automatically
    /// categorizes them using LLM analysis
    /// </summary>
    public class EmailMonitoringService : IDisposable
    {
        private readonly Outlook.Application _outlookApp;
        private readonly List<Outlook.MAPIFolder> _monitoredFolders;
        private readonly List<Outlook.Items> _monitoredFolderItems;
        private bool _isDisposed;

        public EmailMonitoringService(Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp ?? throw new ArgumentNullException(nameof(outlookApp));
            _monitoredFolders = new List<Outlook.MAPIFolder>();
            _monitoredFolderItems = new List<Outlook.Items>();
        }

        /// <summary>
        /// Starts monitoring configured mailboxes
        /// </summary>
        public void StartMonitoring()
        {
            if (!ThisAddIn.userdata.EmailMonitoringEnabled)
                return;

            if (ThisAddIn.userdata.MonitoredMailboxes == null || !ThisAddIn.userdata.MonitoredMailboxes.Any())
                return;

            try
            {
                // Get all stores (mailboxes)
                foreach (Outlook.Store store in _outlookApp.Session.Stores)
                {
                    try
                    {
                        string storeName = store.DisplayName;

                        // Check if this store should be monitored
                        if (ThisAddIn.userdata.MonitoredMailboxes.Contains(storeName))
                        {
                            // Get the Inbox folder of this store
                            Outlook.MAPIFolder rootFolder = store.GetRootFolder();
                            Outlook.MAPIFolder inboxFolder = GetInboxFolder(rootFolder);

                            if (inboxFolder != null)
                            {
                                AttachToFolder(inboxFolder);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log error but continue with other stores
                        System.Diagnostics.Debug.WriteLine($"Error monitoring store: {ex.Message}");
                    }
                    finally
                    {
                        if (store != null)
                            Marshal.ReleaseComObject(store);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error starting email monitoring: {ex.Message}");
            }
        }

        /// <summary>
        /// Stops monitoring all mailboxes and releases resources
        /// </summary>
        public void StopMonitoring()
        {
            DetachAllFolders();
        }

        /// <summary>
        /// Attaches event handlers to a folder to monitor new emails
        /// </summary>
        private void AttachToFolder(Outlook.MAPIFolder folder)
        {
            try
            {
                Outlook.Items items = folder.Items;
                items.ItemAdd += OnNewMailItem;

                _monitoredFolders.Add(folder);
                _monitoredFolderItems.Add(items);

                System.Diagnostics.Debug.WriteLine($"Monitoring started for folder: {folder.Name}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error attaching to folder: {ex.Message}");
            }
        }

        /// <summary>
        /// Detaches event handlers from all monitored folders
        /// </summary>
        private void DetachAllFolders()
        {
            foreach (var items in _monitoredFolderItems)
            {
                try
                {
                    items.ItemAdd -= OnNewMailItem;
                }
                catch { }
            }

            foreach (var folder in _monitoredFolders)
            {
                try
                {
                    Marshal.ReleaseComObject(folder);
                }
                catch { }
            }

            _monitoredFolderItems.Clear();
            _monitoredFolders.Clear();
        }

        /// <summary>
        /// Event handler for new email items
        /// </summary>
        private void OnNewMailItem(object item)
        {
            Outlook.MailItem mailItem = item as Outlook.MailItem;
            if (mailItem == null)
                return;

            // Process email asynchronously to avoid blocking Outlook
            Task.Run(async () =>
            {
                try
                {
                    await ProcessNewEmail(mailItem).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error processing new email: {ex.Message}");
                }
                finally
                {
                    // Release COM object
                    if (mailItem != null)
                        Marshal.ReleaseComObject(mailItem);
                }
            });
        }

        /// <summary>
        /// Processes a new email: classifies it and optionally generates a reply draft
        /// </summary>
        private async Task ProcessNewEmail(Outlook.MailItem mailItem)
        {
            if (mailItem == null)
                return;

            // Get enabled categories
            var enabledCategories = ThisAddIn.userdata.EmailCategories?
                .Where(c => c.IsEnabled)
                .ToList();

            if (enabledCategories == null || !enabledCategories.Any())
                return;

            // Classify the email
            EmailCategory assignedCategory = await ClassifyEmail(mailItem, enabledCategories).ConfigureAwait(false);

            if (assignedCategory != null)
            {
                // Assign Outlook category
                AssignOutlookCategory(mailItem, assignedCategory.CategoryName);

                // Generate reply draft if configured
                if (assignedCategory.GenerateReplyDraft && !string.IsNullOrWhiteSpace(assignedCategory.ReplyPrompt))
                {
                    await GenerateReplyDraft(mailItem, assignedCategory).ConfigureAwait(false);
                }
            }
        }

        /// <summary>
        /// Uses LLM to classify the email into one of the configured categories
        /// </summary>
        private async Task<EmailCategory> ClassifyEmail(Outlook.MailItem mailItem, List<EmailCategory> categories)
        {
            try
            {
                // Extract email information
                string subject = mailItem.Subject ?? "";
                string sender = mailItem.SenderName ?? "";
                string body = GetEmailBodyText(mailItem);

                // Build classification prompt
                string classificationPrompt = BuildClassificationPrompt(subject, sender, body, categories);

                // Call LLM
                string llmResponse = await ThisAddIn.GetLLMResponse(classificationPrompt).ConfigureAwait(false);

                // Parse response to find matching category
                EmailCategory matchedCategory = ParseCategoryFromResponse(llmResponse, categories);

                return matchedCategory;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error classifying email: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Builds the LLM prompt for email classification
        /// </summary>
        private string BuildClassificationPrompt(string subject, string sender, string body, List<EmailCategory> categories)
        {
            var sb = new StringBuilder();

            sb.AppendLine("You are an email classification assistant. Analyze the following email and determine which category it belongs to.");
            sb.AppendLine();
            sb.AppendLine("EMAIL DETAILS:");
            sb.AppendLine($"From: {sender}");
            sb.AppendLine($"Subject: {subject}");
            sb.AppendLine($"Body: {TruncateText(body, 2000)}"); // Limit body length
            sb.AppendLine();
            sb.AppendLine("AVAILABLE CATEGORIES:");

            for (int i = 0; i < categories.Count; i++)
            {
                sb.AppendLine($"\n{i + 1}. Category: {categories[i].CategoryName}");
                sb.AppendLine($"   Classification Rule: {categories[i].ClassificationPrompt}");
                if (!string.IsNullOrWhiteSpace(categories[i].Description))
                {
                    sb.AppendLine($"   Description: {categories[i].Description}");
                }
            }

            sb.AppendLine();
            sb.AppendLine("INSTRUCTIONS:");
            sb.AppendLine("- Analyze the email against each category's classification rule");
            sb.AppendLine("- Choose the MOST appropriate category");
            sb.AppendLine("- If no category matches, respond with 'NONE'");
            sb.AppendLine("- Respond with ONLY the category name (exact match required)");
            sb.AppendLine();
            sb.AppendLine("Your response (category name only):");

            return sb.ToString();
        }

        /// <summary>
        /// Parses the LLM response to extract the category name
        /// </summary>
        private EmailCategory ParseCategoryFromResponse(string llmResponse, List<EmailCategory> categories)
        {
            if (string.IsNullOrWhiteSpace(llmResponse))
                return null;

            string response = llmResponse.Trim().ToLowerInvariant();

            if (response.Contains("none"))
                return null;

            // Try exact match first
            foreach (var category in categories)
            {
                if (response.Equals(category.CategoryName, StringComparison.OrdinalIgnoreCase))
                    return category;
            }

            // Try partial match
            foreach (var category in categories)
            {
                if (response.Contains(category.CategoryName.ToLowerInvariant()))
                    return category;
            }

            return null;
        }

        /// <summary>
        /// Assigns an Outlook category to the email
        /// </summary>
        private void AssignOutlookCategory(Outlook.MailItem mailItem, string categoryName)
        {
            try
            {
                // Ensure category exists in Outlook
                EnsureCategoryExists(categoryName);

                // Assign category to email
                if (string.IsNullOrEmpty(mailItem.Categories))
                {
                    mailItem.Categories = categoryName;
                }
                else if (!mailItem.Categories.Contains(categoryName))
                {
                    mailItem.Categories += ", " + categoryName;
                }

                mailItem.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error assigning category: {ex.Message}");
            }
        }

        /// <summary>
        /// Ensures the category exists in Outlook's category list
        /// </summary>
        private void EnsureCategoryExists(string categoryName)
        {
            try
            {
                Outlook.Categories categories = _outlookApp.Session.Categories;

                bool categoryExists = false;
                foreach (Outlook.Category cat in categories)
                {
                    if (cat.Name.Equals(categoryName, StringComparison.OrdinalIgnoreCase))
                    {
                        categoryExists = true;
                        Marshal.ReleaseComObject(cat);
                        break;
                    }
                    Marshal.ReleaseComObject(cat);
                }

                if (!categoryExists)
                {
                    categories.Add(categoryName, Outlook.OlCategoryColor.olCategoryColorNone);
                }

                Marshal.ReleaseComObject(categories);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error ensuring category exists: {ex.Message}");
            }
        }

        /// <summary>
        /// Generates a reply draft for the email
        /// </summary>
        private async Task GenerateReplyDraft(Outlook.MailItem originalMail, EmailCategory category)
        {
            Outlook.MailItem replyDraft = null;
            try
            {
                // Create reply
                replyDraft = originalMail.Reply() as Outlook.MailItem;

                if (replyDraft == null)
                    return;

                // Build reply generation prompt
                string replyPrompt = BuildReplyPrompt(originalMail, category.ReplyPrompt);

                // Generate reply content using LLM
                string replyContent = await ThisAddIn.GetLLMResponse(replyPrompt).ConfigureAwait(false);

                // Set reply body
                if (!string.IsNullOrWhiteSpace(replyContent))
                {
                    replyDraft.Body = replyContent;
                    replyDraft.Save();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error generating reply draft: {ex.Message}");
            }
            finally
            {
                if (replyDraft != null)
                    Marshal.ReleaseComObject(replyDraft);
            }
        }

        /// <summary>
        /// Builds the prompt for reply generation
        /// </summary>
        private string BuildReplyPrompt(Outlook.MailItem originalMail, string replyPromptTemplate)
        {
            var sb = new StringBuilder();

            sb.AppendLine("Generate a professional email reply based on the following:");
            sb.AppendLine();
            sb.AppendLine("ORIGINAL EMAIL:");
            sb.AppendLine($"From: {originalMail.SenderName}");
            sb.AppendLine($"Subject: {originalMail.Subject}");
            sb.AppendLine($"Body: {TruncateText(GetEmailBodyText(originalMail), 2000)}");
            sb.AppendLine();
            sb.AppendLine("REPLY INSTRUCTIONS:");
            sb.AppendLine(replyPromptTemplate);
            sb.AppendLine();
            sb.AppendLine("Generate the reply email body:");

            return sb.ToString();
        }

        /// <summary>
        /// Gets the inbox folder from a root folder
        /// </summary>
        private Outlook.MAPIFolder GetInboxFolder(Outlook.MAPIFolder rootFolder)
        {
            try
            {
                // Try to get default inbox folder
                foreach (Outlook.MAPIFolder folder in rootFolder.Folders)
                {
                    if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
                    {
                        // Check if it's the inbox (usually has name "Inbox" or localized equivalent)
                        if (folder.Name.ToLowerInvariant().Contains("inbox") ||
                            folder.Name.ToLowerInvariant().Contains("posteingang"))
                        {
                            return folder;
                        }
                    }
                }

                // Fallback: Get first mail folder
                foreach (Outlook.MAPIFolder folder in rootFolder.Folders)
                {
                    if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
                    {
                        return folder;
                    }
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Extracts plain text from email body
        /// </summary>
        private string GetEmailBodyText(Outlook.MailItem mailItem)
        {
            try
            {
                // Try plain text first
                if (!string.IsNullOrWhiteSpace(mailItem.Body))
                    return mailItem.Body;

                // Fallback to HTML body stripped of tags
                if (!string.IsNullOrWhiteSpace(mailItem.HTMLBody))
                {
                    return StripHtmlTags(mailItem.HTMLBody);
                }
            }
            catch { }

            return "";
        }

        /// <summary>
        /// Strips HTML tags from text
        /// </summary>
        private string StripHtmlTags(string html)
        {
            if (string.IsNullOrWhiteSpace(html))
                return "";

            // Basic HTML tag removal
            string text = Regex.Replace(html, @"<[^>]+>", "");
            text = Regex.Replace(text, @"\s+", " ");
            return text.Trim();
        }

        /// <summary>
        /// Truncates text to specified length
        /// </summary>
        private string TruncateText(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;

            return text.Substring(0, maxLength) + "...";
        }

        public void Dispose()
        {
            if (_isDisposed)
                return;

            StopMonitoring();
            _isDisposed = true;
        }
    }
}
