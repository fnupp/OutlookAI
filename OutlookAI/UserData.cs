using System.Collections.Generic;

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

        public bool OpenAIAPIActive { get; set; }
        public string OpenAIAPIUrl { get; set; }
        public string OpenAIAPIKey { get; set; }
        public string OpenAIAPIModel { get; set; }


        public bool OllamaActive { get; set; }
        public string OllamaUrl { get; set; }
        public string Ollamamodel { get; set; }

        public List<string> OllamaModels { get; set; }

    }
}
