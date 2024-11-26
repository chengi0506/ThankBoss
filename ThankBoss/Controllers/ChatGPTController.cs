using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace ThankBoss.Controllers
{
    public class ChatGPT
    {
        public static Result CallChatGPT(string msg)
        {
            HttpClient client = new HttpClient();
            string uri = "https://api.openai.com/v1/chat/completions";

            // Request headers.
            client.DefaultRequestHeaders.Add(
                "Authorization", "Bearer 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");

            object jsonModel = null; // 在這裡宣告 jsonModel

                // 建立 JSON 物件
                jsonModel = new
                {
                    model = "gpt-3.5-turbo",
                    messages = new[]
                    {
                        new { role = "system", content = "我希望你扮演一位客服人員。你是一位友善和樂於助人的人，專門負責回答客戶的任何問題和提供解決方案。你具備出色的溝通和解決問題的能力，能夠快速理解客戶的需求並給予準確的回答。你的目標是確保客戶滿意，並幫助他們解決遇到的所有問題。" },
                        //new { role = "system", content = "我希望你能擔任譽為星座專家的人物，擁有深厚的占星知識和研究經驗。你的專業領域涵蓋十二星座、行星運動、宮位配置等方面，你經常透過占星術的眼光來解讀人們的性格、行為和命運。\r\n\r\n作為星座專家，你具備靈活的洞察力，能夠根據星象的變化預測個人和集體的運勢。你常常受到人們的請求，尋求有關愛情、事業、人際關係等方面的建議。你相信星座可以提供有價值的指引，幫助人們更好地了解自己，走向更加適合他們的道路。\r\n\r\n當人們向我請教時，你會以一種充滿熱情和深度的語氣回答。你用易懂的方式解釋星座的複雜性，幫助他們理解星象對他們生活的影響。你的目標是協助人們通過星座的智慧，找到生活中的平衡點，並鼓勵他們迎接未來的挑戰。" },
                        new { role = "user", content = msg },
                        new { role = "assistant", content = $"" }
                    }
                };

            // 將 JSON 物件轉換成 JSON 字串
            var jsonString = JsonConvert.SerializeObject(jsonModel, Formatting.Indented);

            var content = new StringContent(jsonString, Encoding.UTF8, "application/json");
            var response = client.PostAsync(uri, content).Result;
            var JSON = response.Content.ReadAsStringAsync().Result;

            return Newtonsoft.Json.JsonConvert.DeserializeObject<Result>(JSON);
        }

        static string ReplaceMiddleCharacter(string input, char replacement)
        {
            if (input.Length % 2 == 0 || input.Length < 3)
            {
                // If the length is even or less than 3, it doesn't have a distinct middle character.
                // Handle such cases based on your requirements.
                return input;
            }

            int middleIndex = input.Length / 2;
            string modifiedString = input.Substring(0, middleIndex) + replacement + input.Substring(middleIndex + 1);
            return modifiedString;
        }
    }

    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    public class Choice
    {
        public int index { get; set; }
        public Message message { get; set; }
        public string finish_reason { get; set; }
    }

    public class Message
    {
        public string role { get; set; }
        public string content { get; set; }
    }

    public class Result
    {
        public string id { get; set; }
        public string @object { get; set; }
        public int created { get; set; }
        public List<Choice> choices { get; set; }
        public Usage usage { get; set; }
    }

    public class Usage
    {
        public int prompt_tokens { get; set; }
        public int completion_tokens { get; set; }
        public int total_tokens { get; set; }
    }
}