using System;

namespace OutlookAI
{
    /// <summary>
    /// Represents an email category configuration for automatic classification and reply generation
    /// </summary>
    public class EmailCategory
    {
        /// <summary>
        /// The name of the Outlook category to assign to emails
        /// </summary>
        public string CategoryName { get; set; }

        /// <summary>
        /// The LLM prompt used to determine if an email belongs to this category
        /// This prompt will be included in the classification request along with email content
        /// </summary>
        public string ClassificationPrompt { get; set; }

        /// <summary>
        /// Whether to automatically generate a draft reply for emails in this category
        /// </summary>
        public bool GenerateReplyDraft { get; set; }

        /// <summary>
        /// The LLM prompt used to generate reply content when GenerateReplyDraft is true
        /// Can include §§Input§§ placeholder for additional context
        /// </summary>
        public string ReplyPrompt { get; set; }

        /// <summary>
        /// Whether this category is currently active/enabled
        /// </summary>
        public bool IsEnabled { get; set; }

        /// <summary>
        /// Optional description of the category for user reference
        /// </summary>
        public string Description { get; set; }

        public EmailCategory()
        {
            CategoryName = string.Empty;
            ClassificationPrompt = string.Empty;
            ReplyPrompt = string.Empty;
            Description = string.Empty;
            GenerateReplyDraft = false;
            IsEnabled = true;
        }

        /// <summary>
        /// Creates a deep copy of this EmailCategory
        /// </summary>
        public EmailCategory Clone()
        {
            return new EmailCategory
            {
                CategoryName = this.CategoryName,
                ClassificationPrompt = this.ClassificationPrompt,
                GenerateReplyDraft = this.GenerateReplyDraft,
                ReplyPrompt = this.ReplyPrompt,
                IsEnabled = this.IsEnabled,
                Description = this.Description
            };
        }

        public override string ToString()
        {
            return CategoryName;
        }
    }
}
