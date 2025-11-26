using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAI
{
    public class UserData
    {
        public string Prompt1 { get; set; }
        public string Prompt2 { get; set; }
        public string Prompt3 { get; set; }
        public string Prompt4 { get; set; }
        public string Titel1 { get; set; }
        public string Titel2 { get; set; }
        public string Titel3 { get; set; }
        public string Titel4 { get; set; }


        public string Summary1 { get; set; }
        public string SummaryTitel1 { get; set; }
        public bool? SummaryMultiple1 { get; set; }
        public string Summary2 { get; set; }
        public string SummaryTitel2 { get; set; }
        public bool? SummaryMultiple2 { get; set; }


        public string ComposePrompt1 { get; set; }
        public string ComposeTitle1 { get; set; }
        public string ComposePrompt2 { get; set; }
        public string ComposeTitle2 { get; set; }
        public string ComposePrompt3 { get; set; }
        public string ComposeTitle3 { get; set; }


        public bool OpenAIAPIActive { get; set; }
        public string OpenAIAPIUrl { get; set; }

        // Encrypted API key will be serialized
        public string EncryptedOpenAIAPIKey
        {
            get => encryptedOpenAIAPIKey;
            set => encryptedOpenAIAPIKey = value;
        }

        [JsonIgnore]
        private string encryptedOpenAIAPIKey;

        [IgnoreDataMember]
        public string OpenAIAPIKey
        {
            get
            {
                return string.IsNullOrEmpty(encryptedOpenAIAPIKey)
                    ? null
                    : DecryptPassword(encryptedOpenAIAPIKey);
            }
            set
            {
                encryptedOpenAIAPIKey = string.IsNullOrEmpty(value)
                    ? null
                    : EncryptPassword(value);
            }
        }

        public string OpenAIAPIModel { get; set; }


        public bool OllamaActive { get; set; }
        public string OllamaUrl { get; set; }
        public string Ollamamodel { get; set; }

        public List<string> OllamaModels { get; set; }


        // Email Monitoring and Categorization Settings
        public bool EmailMonitoringEnabled { get; set; }
        public List<string> MonitoredMailboxes { get; set; }
        public List<EmailCategory> EmailCategories { get; set; }


        public bool ProxyActive { get; set; }
        public string ProxyUrl { get; set; }
        public string ProxyUsername { get; set; }

        // Verschlüsseltes Passwort wird serialisiert
        public string EncryptedProxyPassword
        {
            get => encryptedProxyPassword;
            set => encryptedProxyPassword = value;
        }

        [NonSerialized]
        private string encryptedProxyPassword;

        [IgnoreDataMember]
        public string ProxyPassword
        {
            get
            {
                return string.IsNullOrEmpty(encryptedProxyPassword)
                    ? null
                    : DecryptPassword(encryptedProxyPassword);
            }
            set
            {
                encryptedProxyPassword = string.IsNullOrEmpty(value)
                    ? null
                    : EncryptPassword(value);
            }
        }

        private string EncryptPassword(string password)
        {
            var data = Encoding.UTF8.GetBytes(password);
            var encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(encrypted);
        }

        private string DecryptPassword(string encryptedPassword)
        {
            var data = Convert.FromBase64String(encryptedPassword);
            var decrypted = ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(decrypted);
        }
    }
}