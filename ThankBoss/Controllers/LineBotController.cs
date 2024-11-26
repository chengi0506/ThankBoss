using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Http;

using isRock.LineBot;
using System.Drawing;
using System.EnterpriseServices;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using System.Text;
using System.Net.Http;
using System.Threading.Tasks;
using System.Configuration;
using WebGrease;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Data.SqlClient;
using System.Data;
using System.Reflection;
using System.Security.Policy;
using System.Web.DynamicData;
using System.Net;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

using Tesseract;
using System.Drawing.Drawing2D;

namespace ThankBoss.Controllers
{
    public class LineBotController : isRock.LineBot.LineWebHookControllerBase
    {
        private readonly string channelAccessToken;
        private readonly string AdminUserId;

        enum Command
        {
            指令,
            涼涼名單,
            查詢活動,
            參加名單,
            不在勤名單,
            加班名單,
            筆電掃描,
            謝謝老闆
        }

        public LineBotController()
        {
            channelAccessToken = Properties.Settings.Default.ChannelAccessToken;
            AdminUserId = Properties.Settings.Default.AdminUserId;
        }

        [Route("api/LineBotWebHook")]
        [HttpPost]
        public async Task<IHttpActionResult> POSTAsync()
        {
            var AdminUserId = this.AdminUserId;

            try
            {
                //設定ChannelAccessToken
                this.ChannelAccessToken = this.channelAccessToken;
                //配合Line Verify
                if (ReceivedMessage.events == null || ReceivedMessage.events.Count() <= 0 ||
                    ReceivedMessage.events.FirstOrDefault().replyToken == "00000000000000000000000000000000") return Ok();
                //取得Line Event
                var LineEvent = this.ReceivedMessage.events.FirstOrDefault();
                //準備回覆訊息
                if (LineEvent.type.ToLower() == "message" && LineEvent.message.type == "text")
                {
                    if (LineEvent.message.text == Command.指令.ToString())
                    {
                        try
                        {
                            // 產生指令明細集
                            var details = Enum.GetValues(typeof(Command))
                                              .Cast<Command>()
                                              .Select(command => new
                                              {
                                                  type = "button",
                                                  style = "link",
                                                  height = "sm",
                                                  action = new
                                                  {
                                                      type = "message",
                                                      label = command.ToString(),
                                                      text = command.ToString()
                                                  }
                                              })
                                              .ToList();

                            JObject footer = new JObject(
                                new JProperty("type", "box"),
                                new JProperty("layout", "vertical"),
                                new JProperty("spacing", "sm"),
                                new JProperty("flex", 0),
                                new JProperty("contents", JArray.FromObject(details))
                            );

                            JObject body = new JObject(
                                new JProperty("type", "box"),
                                new JProperty("layout", "vertical"),
                                new JProperty("contents", new JArray(
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("text", "所有指令："),
                                        new JProperty("weight", "bold"),
                                        new JProperty("size", "xl")
                                    )
                                ))
                            );

                            JObject output = new JObject(
                                new JProperty("type", "bubble"),
                                new JProperty("body", body),
                                new JProperty("footer", footer)
                            );

                            JObject contents = new JObject(
                                new JProperty("type", "flex"),
                                new JProperty("altText", "所有指令："),
                                new JProperty("contents", output)
                            );


                            // 將結果轉換成 JSON 字串並輸出
                            string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);

                            // 將結果輸出至 output.json 檔案
                            System.IO.File.WriteAllText(HttpContext.Current.Server.MapPath("~/Activity/output.json"), jsonOutput);

                            //回覆訊息
                            ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);
                        }
                        catch (Exception ex)
                        {
                            //回覆訊息
                            //ReplyMessage(LineEvent.replyToken, ex.Message);
                            PushMessagesWithJSON(AdminUserId, ex.Message);
                        }
                        
                    }
                    // 判斷訊息是否包含 "涼涼名單" 關鍵字
                    if (LineEvent.message.text.Contains(Command.涼涼名單.ToString()))
                    {
                        //groupId
                        string groupId = LineEvent.source.groupId;
                        //userId
                        string userId = LineEvent.source.userId;


                        // 呼叫 GetUserInfo 方法以獲取使用者名稱和大頭貼
                        var userInfo = await GetUserInfo(userId);

                        // 從返回的元組中獲取使用者名稱和大頭貼 URL
                        string userName = userInfo.Item1;
                        string profilePictureUrl = userInfo.Item2;


                        // 讀取輸入的 CSV 檔案
                        //string[] lines = File.ReadAllLines(HttpContext.Current.Server.MapPath("~/Order/input.csv"), Encoding.GetEncoding("UTF-8"));

                        //訂單檔路徑
                        string filePath = HttpContext.Current.Server.MapPath("~/Order/input.csv");
                        
                        //string filePath = @"\\FCLDWEB\PosVersionService\order\input.csv";
                        
                        // 創建一個動態陣列來存儲檔案內容
                        List<string> linesList = new List<string>();

                        // 使用 FileStream 打開檔案，允許其他進程讀取檔案
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            // 使用 StreamReader 讀取檔案內容
                            using (StreamReader reader = new StreamReader(fs, Encoding.GetEncoding("UTF-8")))
                            {
                                // 逐行讀取檔案內容
                                string line;
                                while ((line = reader.ReadLine()) != null)
                                {
                                    // 將每一行的內容添加到動態陣列中
                                    linesList.Add(line);
                                }
                            }
                        }

                        // 將動態陣列轉換為 string[] 陣列
                        string[] lines = linesList.ToArray();
                        //商店名稱
                        string store = lines[0].Split(',')[0];
                        //收款人姓名
                        string Payee = lines[0].Split(',')[8];
                        //預估抵達時間
                        string activeOrderStatus = lines[0].Split(',')[9];
                        //未收人數
                        int unpaidCount = 0;

                        // contents
                        List<JObject> contentsList = new List<JObject>();

                        //結束列數
                        int endIndex = -1;

                        // 尋找第一欄包含關鍵字"總原價"的索引
                        for (int i = 0; i < lines.Length; i++)
                        {
                            string[] columns = lines[i].Split(',');

                            if (columns.Length > 0 && columns[0] == "總原價")
                            {
                                endIndex = i;
                                break;
                            }
                        }

                        //總應付
                        string totalPayment = lines[endIndex].Split(',')[5];

                        for (int i = 1; i < endIndex; i++)
                        {
                            string[] columns = lines[i].Split(',');

                            //如果已收款標誌為非V，則計算未收款人數
                            if (columns[6].Trim() != "V")
                            {
                                unpaidCount++;
                            }

                            // 為 CSV 每一行構造 JSON 物件
                            JObject contentsObj = new JObject(
                                new JProperty("type", "box"),
                                new JProperty("layout", "horizontal"),
                                new JProperty("contents",
                                    new JArray(
                                        //new JObject(
                                        //    new JProperty("type", "text"),
                                        //    new JProperty("text", columns[0]),
                                        //    new JProperty("size", "sm"),
                                        //    new JProperty("color", "#1167b1"),
                                        //    new JProperty("weight", "bold")
                                        //),

                                        new JObject(
                                            new JProperty("type", "button"),
                                            new JProperty("color", "#1167b1"),
                                            new JProperty("action",
                                                new JObject(
                                                    new JProperty("type", "postback"),  // 將操作類型更改為 postback
                                                    new JProperty("label", columns[0]),
                                                    new JProperty("data", columns[0])
                                                )
                                            ),
                                            new JProperty("style", "link"),
                                            new JProperty("offsetBottom", "10px"),
                                            new JProperty("flex", 3)
                                         ),

                                        new JObject(
                                            new JProperty("type", "text"),
                                            new JProperty("text", $"{columns[5].ToString().Replace(".00", "")}元"),
                                            new JProperty("align", "center"),
                                            new JProperty("size", "md"),
                                            new JProperty("color", "#040420"),
                                            new JProperty("weight", "bold"),
                                            new JProperty("flex", 2)
                                        ),
                                        new JObject(
                                            new JProperty("type", "text"),
                                            new JProperty("text", columns[6].Trim() == "V" ? "已收款" : "未收款"),
                                            new JProperty("align", "center"),
                                            new JProperty("size", "md"),
                                            new JProperty("color", columns[6].Trim() == "V" ? "#040420" : "#9e0200"),
                                            new JProperty("flex", 2)
                                        )
                                    )
                                ),
                                new JProperty("margin", "none")
                            );

                            contentsList.Add(contentsObj);
                        }

                        // 輸出準備 JSON 結構
                        JObject bodyContents = new JObject(
                            new JProperty("type", "box"),
                            new JProperty("layout", "vertical"),
                            new JProperty("contents",
                                new JArray(
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("text", $"{DateTime.Now.ToString("yyyy/MM/dd")}({ConvertToChineseDayOfWeek(DateTime.Now.DayOfWeek)})"),
                                        new JProperty("size", "md"),
                                        new JProperty("color", "#040420"),
                                        new JProperty("weight", "bold")
                                    ),
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("text", $"【{store}】"),
                                        new JProperty("size", "md"),
                                        new JProperty("color", "#040420"),
                                        new JProperty("weight", "bold")
                                    ),
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("text", $"{activeOrderStatus}"),
                                        new JProperty("size", "md"),
                                        new JProperty("color", "#040420"),
                                        new JProperty("weight", "bold")
                                    ),
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("text", $"收款人:{Payee}"),
                                        new JProperty("color", "#9e0200"),
                                        new JProperty("weight", "bold"),
                                        new JProperty("size", "md"),
                                        new JProperty("align", "end")
                                    ),
                                    new JObject(
                                        new JProperty("type", "separator")
                                    ),
                                    new JObject(
                                        new JProperty("type", "box"),
                                        new JProperty("layout", "vertical"),
                                        new JProperty("contents", new JArray(contentsList))
                                    ),
                                    new JObject(
                                        new JProperty("type", "separator")
                                    )
                                )
                            ),
                            new JProperty("paddingStart", "xs"),
                            new JProperty("paddingEnd", "xs")
                        );


                        string uri = $"";

                        JObject footerContents = new JObject(
                        new JProperty("type", "box"),
                        new JProperty("layout", "vertical"),
                        new JProperty("contents", new JArray(
                            new JObject(
                                new JProperty("type", "box"),
                                new JProperty("layout", "horizontal"),
                                new JProperty("contents", new JArray(
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("text", $"總金額:{totalPayment}元"),
                                        new JProperty("weight", "bold"),
                                        new JProperty("size", "md")
                                    ),
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("align", "end"),
                                        new JProperty("text", unpaidCount == 0 ? "收款完成" : $"未收款:{unpaidCount}人"),
                                        new JProperty("color", unpaidCount == 0 ? "#1167b1" : "#9e0200"),
                                        new JProperty("weight", "bold"),
                                        new JProperty("size", "md")
                                    )
                                ))
                            ),
                            new JObject(
                                new JProperty("type", "box"),
                                new JProperty("layout", "vertical"),
                                new JProperty("contents", new JArray(
                                    new JObject(new JProperty("type", "filler")),
                                    new JObject(new JProperty("type", "separator")),
                                    new JObject(
                                        new JProperty("type", "button"),
                                        new JProperty("style", "primary"),
                                        new JProperty("action",
                                        new JObject(
                                            new JProperty("type", "uri"),
                                            new JProperty("label", "LINE Pay"),
                                            new JProperty("uri", uri)
                                        ))
                                    )
                                )),
                                new JProperty("spacing", "sm")
                            )
                        ))
                    );


                        JObject body = new JObject(
                            new JProperty("type", "bubble"),
                            new JProperty("hero", new JObject(
                                new JProperty("type", "image"),
                                new JProperty("url", ""),
                                new JProperty("size", "xxl"),
                                new JProperty("aspectRatio", "50:50")
                            )),
                            new JProperty("body", bodyContents),
                            new JProperty("footer", footerContents)
                        );

                        JObject contents = new JObject(
                            new JProperty("type", "flex"),
                            new JProperty("altText", "涼涼名單"),
                            new JProperty("contents", body)
                        );


                        // 將結果轉換成 JSON 字串並輸出
                        string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);

                        // 將結果輸出至 output.json 檔案
                        System.IO.File.WriteAllText(HttpContext.Current.Server.MapPath("~/Order/output.json"), jsonOutput);


                        //回覆訊息
                        ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);


                        //PushMessagesWithJSON(AdminUserId, jsonString);
                    }

                    // 判斷訊息是否包含 "查詢活動" 關鍵字
                    if (LineEvent.message.text == Command.查詢活動.ToString())
                    {
                        //groupId
                        string groupId = LineEvent.source.groupId;
                        //userId
                        string userId = LineEvent.source.userId;


                        // 呼叫 GetUserInfo 方法以獲取使用者名稱和大頭貼
                        var userInfo = await GetUserInfo(userId);

                        // 從返回的元組中獲取使用者名稱和大頭貼 URL
                        string userName = userInfo.Item1;
                        string profilePictureUrl = userInfo.Item2;

                        // 訂單檔路徑
                        string filePath = HttpContext.Current.Server.MapPath("~/Activity/input.csv");
                        // 創建一個動態陣列來存儲檔案內容
                        List<string> linesList = new List<string>();

                        // 使用 FileStream 打開檔案，允許其他進程讀取檔案
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            // 使用 StreamReader 讀取檔案內容
                            using (StreamReader reader = new StreamReader(fs, Encoding.GetEncoding("UTF-8")))
                            {
                                // 逐行讀取檔案內容
                                string line;
                                while ((line = reader.ReadLine()) != null)
                                {
                                    // 將每一行的內容添加到動態陣列中
                                    linesList.Add(line);
                                }
                            }
                        }

                        // 將動態陣列轉換為 string[] 陣列
                        string[] lines = linesList.ToArray();
                        //活動名稱
                        string AName = lines[0].Split(',')[1];
                        //活動日期
                        string ADate = lines[0].Split(',')[3];
                        //活動時間
                        string ATime = lines[0].Split(',')[5];
                        //活動地點
                        string ALocal = lines[0].Split(',')[7];

                        // contents
                        List<JObject> contentsList = new List<JObject>();

                        // 解析 CSV 內容並建立 JSON 物件
                        JObject contentsObj = new JObject(
                            new JProperty("type", "bubble"),
                            new JProperty("hero",
                                new JObject(
                                    new JProperty("type", "image"),
                                    new JProperty("url", ""),
                                    new JProperty("size", "full"),
                                    new JProperty("aspectRatio", "20:13"),
                                    new JProperty("aspectMode", "cover"),
                                    new JProperty("action",
                                        new JObject(
                                            new JProperty("type", "uri"),
                                            new JProperty("uri", "http://linecorp.com/")
                                        )
                                    )
                                )
                            ),
                            new JProperty("body",
                                new JObject(
                                    new JProperty("type", "box"),
                                    new JProperty("layout", "vertical"),
                                    new JProperty("spacing", "md"),
                                    new JProperty("contents",
                                        new JArray(
                                            new JObject(
                                                new JProperty("type", "text"),
                                                new JProperty("text", AName),
                                                new JProperty("wrap", true),
                                                new JProperty("weight", "bold"),
                                                new JProperty("gravity", "center"),
                                                new JProperty("size", "xl")
                                            ),
                                            new JObject(
                                                new JProperty("type", "box"),
                                                new JProperty("layout", "vertical"),
                                                new JProperty("margin", "lg"),
                                                new JProperty("spacing", "sm"),
                                                new JProperty("contents",
                                                    new JArray(
                                                        new JObject(
                                                            new JProperty("type", "box"),
                                                            new JProperty("layout", "baseline"),
                                                            new JProperty("spacing", "sm"),
                                                            new JProperty("contents",
                                                                new JArray(
                                                                    new JObject(
                                                                        new JProperty("type", "text"),
                                                                        new JProperty("text", "日期"),
                                                                        new JProperty("color", "#aaaaaa"),
                                                                        new JProperty("size", "sm"),
                                                                        new JProperty("flex", 1)
                                                                    ),
                                                                    new JObject(
                                                                        new JProperty("type", "text"),
                                                                        new JProperty("text", ADate),
                                                                        new JProperty("wrap", true),
                                                                        new JProperty("size", "sm"),
                                                                        new JProperty("color", "#666666"),
                                                                        new JProperty("flex", 4)
                                                                    )
                                                                )
                                                            )
                                                        ),
                                                        new JObject(
                                                            new JProperty("type", "box"),
                                                            new JProperty("layout", "baseline"),
                                                            new JProperty("spacing", "sm"),
                                                            new JProperty("contents",
                                                                new JArray(
                                                                    new JObject(
                                                                        new JProperty("type", "text"),
                                                                        new JProperty("text", "時間"),
                                                                        new JProperty("color", "#aaaaaa"),
                                                                        new JProperty("size", "sm"),
                                                                        new JProperty("flex", 1)
                                                                    ),
                                                                    new JObject(
                                                                        new JProperty("type", "text"),
                                                                        new JProperty("text", ATime),
                                                                        new JProperty("wrap", true),
                                                                        new JProperty("color", "#666666"),
                                                                        new JProperty("size", "sm"),
                                                                        new JProperty("flex", 4)
                                                                    )
                                                                )
                                                            )
                                                        ),
                                                        new JObject(
                                                            new JProperty("type", "box"),
                                                            new JProperty("layout", "baseline"),
                                                            new JProperty("spacing", "sm"),
                                                            new JProperty("contents",
                                                                new JArray(
                                                                    new JObject(
                                                                        new JProperty("type", "text"),
                                                                        new JProperty("text", "地點"),
                                                                        new JProperty("color", "#aaaaaa"),
                                                                        new JProperty("size", "sm"),
                                                                        new JProperty("flex", 1)
                                                                    ),
                                                                    new JObject(
                                                                        new JProperty("type", "text"),
                                                                        new JProperty("text", ALocal),
                                                                        new JProperty("wrap", true),
                                                                        new JProperty("color", "#666666"),
                                                                        new JProperty("size", "sm"),
                                                                        new JProperty("flex", 4)
                                                                    )
                                                                )
                                                            )
                                                        )
                                                    )
                                                )
                                            ),
                                            new JObject(
                                                new JProperty("type", "box"),
                                                new JProperty("layout", "vertical"),
                                                new JProperty("margin", "xxl"),
                                                new JProperty("contents",
                                                    new JArray(
                                                        new JObject(
                                                            new JProperty("type", "separator")
                                                        ),
                                                        new JObject(
                                                            new JProperty("type", "button"),
                                                            new JProperty("action",
                                                                new JObject(
                                                                    //new JProperty("type", "uri"),
                                                                    //new JProperty("label", "參加"),
                                                                    //new JProperty("uri", $"{uri}&isJoin={HttpUtility.UrlEncode("參加")}")

                                                                    new JProperty("type", "postback"),  // 將操作類型更改為 postback
                                                                    new JProperty("label", "參加"),
                                                                    new JProperty("data", "參加")  // 在 data 中指定用於區分參加的標識
                                                                )
                                                            ),
                                                            new JProperty("style", "link")
                                                        ),
                                                        new JObject(
                                                            new JProperty("type", "separator")
                                                        ),
                                                        new JObject(
                                                            new JProperty("type", "button"),
                                                            new JProperty("action",
                                                                new JObject(
                                                                    //new JProperty("type", "uri"),
                                                                    //new JProperty("label", "不參加"),
                                                                    //new JProperty("uri", $"{uri}&isJoin={HttpUtility.UrlEncode("不參加")}")

                                                                    new JProperty("type", "postback"),  // 將操作類型更改為 postback
                                                                    new JProperty("label", "不參加"),
                                                                    new JProperty("data", "不參加")  // 在 data 中指定用於區分不參加的標識
                                                                )
                                                            ),
                                                            new JProperty("style", "link")
                                                        ),
                                                        new JObject(
                                                            new JProperty("type", "separator")
                                                        ),
                                                        new JObject(
                                                            new JProperty("type", "button"),
                                                            new JProperty("action",
                                                                new JObject(
                                                                    new JProperty("type", "message"),  //  將操作類型更改為 postback
                                                                    new JProperty("label", "參加名單"),
                                                                    new JProperty("text", "參加名單")  // 在 text 中指定用於區分參加名單的標識
                                                                )
                                                            ),
                                                            new JProperty("style", "link")
                                                        )
                                                    )
                                                )
                                            )
                                        )
                                    )
                                )
                            )
                        );

                        JObject contents = new JObject(
                            new JProperty("type", "flex"),
                            new JProperty("altText", "我再補給你"),
                            new JProperty("contents", contentsObj)
                        );

                        // 將結果轉換成 JSON 字串並輸出
                        string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);

                        // 將結果輸出至 output.json 檔案
                        System.IO.File.WriteAllText(HttpContext.Current.Server.MapPath("~/Activity/output.json"), jsonOutput);

                        //回覆訊息
                        ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);
                    }

                    // 判斷訊息是否包含 "參加名單" 關鍵字
                    if (LineEvent.message.text == Command.參加名單.ToString())
                    {
                        //訂單檔路徑
                        string filePath = HttpContext.Current.Server.MapPath("~/Activity/input.csv");
                        // 創建一個動態陣列來存儲檔案內容
                        List<string> linesList = new List<string>();

                        // 使用 FileStream 打開文件，允許其他進程讀取文件
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            // 使用 StreamReader 讀取文件內容
                            using (StreamReader reader = new StreamReader(fs, Encoding.GetEncoding("UTF-8")))
                            {
                                // 逐行讀取文件內容
                                string line;
                                while ((line = reader.ReadLine()) != null)
                                {
                                    // 將每一行的內容添加到動態陣列中
                                    linesList.Add(line);
                                }
                            }
                        }

                        // 將動態陣列轉換為 string[] 陣列
                        string[] lines = linesList.ToArray();
                        //活動名稱
                        string AName = lines[0].Split(',')[1];
                        //活動日期
                        string ADate = lines[0].Split(',')[3];
                        //活動時間
                        string ATime = lines[0].Split(',')[5];
                        //活動地點
                        string ALocal = lines[0].Split(',')[7];
                        //活動人數
                        int Num = 0;


                        // contents
                        List<JObject> contentsList = new List<JObject>();

                        // 解析 CSV 数据並建立訂單明細
                        for (int i = 1; i < lines.Length; i++)
                        {
                            string[] columns = lines[i].Split(',');

                            if (columns[2] == "不參加")
                            {
                                continue;
                            }

                            JObject contentsObj = new JObject(
                                new JProperty("type", "box"),
                                new JProperty("layout", "vertical"),
                                new JProperty("contents",
                                    new JArray(
                                        new JObject(
                                            new JProperty("type", "text"),
                                            new JProperty("text", columns[1]),
                                            new JProperty("size", "md"),
                                            new JProperty("color", "#1167b1"),
                                            new JProperty("weight", "bold")
                                        )
                                    )
                                ),
                                new JProperty("margin", "10px")
                            );

                            Num++;
                            contentsList.Add(contentsObj);
                        }

                        // 創建 JSON 物件
                        JObject body = new JObject(
                            new JProperty("type", "bubble"),
                            new JProperty("hero", new JObject(
                                new JProperty("type", "image"),
                                new JProperty("url", ""),
                                new JProperty("size", "full"),
                                new JProperty("aspectRatio", "20:13"),
                                new JProperty("aspectMode", "cover"),
                                new JProperty("action", new JObject(
                                    new JProperty("type", "uri"),
                                    new JProperty("uri", "http://linecorp.com/")
                                ))
                            )),
                            new JProperty("body", new JObject(
                                new JProperty("type", "box"),
                                new JProperty("layout", "vertical"),
                                new JProperty("spacing", "md"),
                                new JProperty("contents", new JArray(
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("text", AName),
                                        new JProperty("wrap", true),
                                        new JProperty("weight", "bold"),
                                        new JProperty("gravity", "center"),
                                        new JProperty("size", "xl")
                                    ),
                                    new JObject(
                                        new JProperty("type", "box"),
                                        new JProperty("layout", "vertical"),
                                        new JProperty("margin", "lg"),
                                        new JProperty("spacing", "sm"),
                                        new JProperty("contents", new JArray(
                                            new JObject(
                                                new JProperty("type", "box"),
                                                new JProperty("layout", "baseline"),
                                                new JProperty("spacing", "sm"),
                                                new JProperty("contents", new JArray(
                                                    new JObject(
                                                        new JProperty("type", "text"),
                                                        new JProperty("text", "日期"),
                                                        new JProperty("color", "#aaaaaa"),
                                                        new JProperty("size", "sm"),
                                                        new JProperty("flex", 1)
                                                    ),
                                                    new JObject(
                                                        new JProperty("type", "text"),
                                                        new JProperty("text", ADate),
                                                        new JProperty("wrap", true),
                                                        new JProperty("size", "sm"),
                                                        new JProperty("color", "#666666"),
                                                        new JProperty("flex", 4)
                                                    )
                                                ))
                                            ),
                                            new JObject(
                                                new JProperty("type", "box"),
                                                new JProperty("layout", "baseline"),
                                                new JProperty("spacing", "sm"),
                                                new JProperty("contents", new JArray(
                                                    new JObject(
                                                        new JProperty("type", "text"),
                                                        new JProperty("text", "時間"),
                                                        new JProperty("color", "#aaaaaa"),
                                                        new JProperty("size", "sm"),
                                                        new JProperty("flex", 1)
                                                    ),
                                                    new JObject(
                                                        new JProperty("type", "text"),
                                                        new JProperty("text", ATime),
                                                        new JProperty("wrap", true),
                                                        new JProperty("color", "#666666"),
                                                        new JProperty("size", "sm"),
                                                        new JProperty("flex", 4)
                                                    )
                                                ))
                                            ),
                                            new JObject(
                                                new JProperty("type", "box"),
                                                new JProperty("layout", "baseline"),
                                                new JProperty("spacing", "sm"),
                                                new JProperty("contents", new JArray(
                                                    new JObject(
                                                        new JProperty("type", "text"),
                                                        new JProperty("text", "地點"),
                                                        new JProperty("color", "#aaaaaa"),
                                                        new JProperty("size", "sm"),
                                                        new JProperty("flex", 1)
                                                    ),
                                                    new JObject(
                                                        new JProperty("type", "text"),
                                                        new JProperty("text", ALocal),
                                                        new JProperty("wrap", true),
                                                        new JProperty("color", "#666666"),
                                                        new JProperty("size", "sm"),
                                                        new JProperty("flex", 4)
                                                    )
                                                ))
                                            ),
                                            new JObject(
                                                new JProperty("type", "box"),
                                                new JProperty("layout", "baseline"),
                                                new JProperty("spacing", "sm"),
                                                new JProperty("contents", new JArray(
                                                    new JObject(
                                                        new JProperty("type", "text"),
                                                        new JProperty("text", "人數"),
                                                        new JProperty("color", "#aaaaaa"),
                                                        new JProperty("size", "sm"),
                                                        new JProperty("flex", 1)
                                                    ),
                                                    new JObject(
                                                        new JProperty("type", "text"),
                                                        new JProperty("text", $"{Num}人"),
                                                        new JProperty("wrap", true),
                                                        new JProperty("color", "#9e0200"),
                                                        new JProperty("size", "xxl"),
                                                        new JProperty("flex", 4)
                                                    )
                                                ))
                                            )
                                        ))
                                    ),
                                    new JObject(
                                        new JProperty("type", "box"),
                                        new JProperty("layout", "vertical"),
                                        new JProperty("margin", "xxl"),
                                        new JProperty("contents", new JArray(
                                            new JObject(
                                                new JProperty("type", "separator")
                                            ),
                                            new JObject(
                                                new JProperty("type", "box"),
                                                new JProperty("layout", "vertical"),
                                                new JProperty("contents", new JArray(contentsList))
                                            )
                                        ))
                                    )
                                ))
                            ))
                        );


                        JObject contents = new JObject(
                            new JProperty("type", "flex"),
                            new JProperty("altText", "參加名單"),
                            new JProperty("contents", body)
                        );

                        // 將結果轉換成 JSON 字串並輸出
                        string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);

                        // 將結果輸出至 output.json 檔案
                        System.IO.File.WriteAllText(HttpContext.Current.Server.MapPath("~/Activity/output.json"), jsonOutput);

                        //回覆訊息
                        ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);
                    }
                    // 判斷訊息是否包含 "謝謝老闆" 關鍵字
                    if (LineEvent.message.text.Length >= 4 && LineEvent.message.text.Substring(0, 4) == Command.謝謝老闆.ToString())
                    {
                        // 使用 ChatGPT 取得回覆訊息
                        var GptResult = ChatGPT.CallChatGPT(LineEvent.message.text.Replace(Command.謝謝老闆.ToString(), string.Empty)).choices[0].message.content;
                        var responseMsg = $"{GptResult}";
                        // 回覆訊息
                        this.ReplyMessage(LineEvent.replyToken, responseMsg);
                    }

                    // 判斷訊息是否包含 "不在勤名單" 關鍵字
                    if (LineEvent.message.text.Contains(Command.不在勤名單.ToString()))
                    {
                        // 日期
                        DateTime workDate;
                        string formattedDate = string.Empty;
                        try
                        {
                            if (DateTime.TryParseExact(LineEvent.message.text.Substring(0, 10), "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out workDate))
                            {
                                formattedDate = workDate.ToString("yyyy/MM/dd");
                            }
                        }
                        catch
                        {
                            formattedDate = DateTime.Now.ToString("yyyy/MM/dd");
                        }

                        // 將所有指令名稱組合成一個回覆訊息
                        //string replyMessage = $"{formattedDate}不在勤名單：\n";
                        //replyMessage += GetVacationListFromDatabase(formattedDate) + "\n";



                        // 取得不在勤名單
                        string data = GetVacationListFromDatabase(formattedDate);

                        // 產生不在勤名單明細
                        var details = new List<object>();
                        if (string.IsNullOrEmpty(data))
                        {
                            details.Add(new
                            {
                                type = "text",
                                text = "本日沒有不在勤名單~",
                                margin = "md",
                                color = "#040420",
                                weight = "bold",
                            });
                        }
                        else
                        {
                            foreach (var item in data.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                details.Add(new
                                {
                                    type = "text",
                                    text = item.Trim(),
                                    margin = "md",
                                    color = "#040420",
                                });
                            }
                        }

                        // 构建 body 的 JObject
                        JObject body = new JObject(
                            new JProperty("type", "box"),
                            new JProperty("layout", "vertical"),
                            new JProperty("contents", new JArray(
                                new JObject(
                                    new JProperty("type", "text"),
                                    new JProperty("text", $"{formattedDate}不在勤名單：\n"),
                                    new JProperty("weight", "bold"),
                                    new JProperty("size", "xl")
                                ),
                                new JObject(
                                    new JProperty("type", "separator")
                                ),
                                new JObject(
                                    new JProperty("type", "box"),
                                    new JProperty("layout", "vertical"),
                                    new JProperty("contents", JArray.FromObject(details))
                                )
                            ))
                        );

                        // 构建 output 的 JObject
                        JObject output = new JObject(
                            new JProperty("type", "bubble"),
                            new JProperty("hero", new JObject(
                                new JProperty("type", "image"),
                                new JProperty("url", ""),
                                new JProperty("size", "xxl"),
                                new JProperty("aspectRatio", "50:50")
                            )),
                            new JProperty("body", body)
                        );

                        // 构建最终的 JObject
                        JObject contents = new JObject(
                            new JProperty("type", "flex"),
                            new JProperty("altText", "不在勤名單"),
                            new JProperty("contents", output)
                        );

                        // 將結果轉換成 JSON 字串並輸出
                        string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);

                        // 將結果輸出至 output.json 檔案
                        System.IO.File.WriteAllText(HttpContext.Current.Server.MapPath("~/Activity/output.json"), jsonOutput);

                        //回覆訊息
                        ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);
                    }


                    // 判斷訊息是否包含 "加班名單" 關鍵字
                    if (LineEvent.message.text.Contains(Command.加班名單.ToString()))
                    {
                            // 日期
                            DateTime workDate;
                            string formattedDate = string.Empty;
                            try
                            {
                                if (DateTime.TryParseExact(LineEvent.message.text.Substring(0, 10), "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out workDate))
                                {

                                    formattedDate = workDate.ToString("yyyy/MM/dd");
                                }
                            }
                            catch
                            {
                                formattedDate = DateTime.Now.ToString("yyyy/MM/dd");
                            }


                        // 取得加班名單
                        string data = GetOverTimeListFromDatabase(formattedDate);

                            // 產生加班名單明細
                            var details = new List<object>();
                            if (string.IsNullOrEmpty(data))
                            {
                                details.Add(new
                                {
                                    type = "text",
                                    text = "本日沒有加班名單~",
                                    margin = "md",
                                    color = "#040420",
                                    weight = "bold",
                                });
                            }
                            else
                            {
                                foreach (var item in data.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries))
                                {
                                    details.Add(new
                                    {
                                        type = "text",
                                        text = item.Trim(),
                                        margin = "md",
                                        color = "#040420",
                                        weight = "regular",
                                    });
                                }
                            }

                            // 构建 body 的 JObject
                            JObject body = new JObject(
                                new JProperty("type", "box"),
                                new JProperty("layout", "vertical"),
                                new JProperty("contents", new JArray(
                                    new JObject(
                                        new JProperty("type", "text"),
                                        new JProperty("text", $"{formattedDate}加班名單：\n"),
                                        new JProperty("weight", "bold"),
                                        new JProperty("size", "xl")
                                    ),
                                    new JObject(
                                        new JProperty("type", "separator")
                                    ),
                                    new JObject(
                                        new JProperty("type", "box"),
                                        new JProperty("layout", "vertical"),
                                        new JProperty("contents", JArray.FromObject(details))
                                    )
                                ))
                            );

                            // 构建 output 的 JObject
                            JObject output = new JObject(
                                new JProperty("type", "bubble"),
                                new JProperty("hero", new JObject(
                                    new JProperty("type", "image"),
                                    new JProperty("url", ""),
                                    new JProperty("size", "xxl"),
                                    new JProperty("aspectRatio", "50:50")
                                )),
                                new JProperty("body", body)
                            );

                            // 构建最终的 JObject
                            JObject contents = new JObject(
                                new JProperty("type", "flex"),
                                new JProperty("altText", "加班名單"),
                                new JProperty("contents", output)
                            );

                            // 將結果轉換成 JSON 字串並輸出
                            string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);

                            // 將結果輸出至 output.json 檔案
                            System.IO.File.WriteAllText(HttpContext.Current.Server.MapPath("~/Activity/output.json"), jsonOutput);

                            //回覆訊息
                            ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);
                    }

                    // 判斷訊息是否包含 "筆電掃描" 關鍵字
                    if (LineEvent.message.text.Contains(Command.筆電掃描.ToString()))
                    {
                        //如果訊息開頭不為"NB"且長度小於8
                        if (!LineEvent.message.text.StartsWith("NB") && LineEvent.message.text.Length < 8)
                        {
                            return Ok();
                        }

                        string NB = LineEvent.message.text.Substring(0, 4);

                        // 打開連接並查詢
                        using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["NBScanDbContext"].ConnectionString))
                        {
                            connection.Open();

                            // UPDATE SQL語句，將 ScanTime 設定為當前時間
                            string sql = "UPDATE News SET ScanTime = @ScanTime,NewsImg = @NewsImg WHERE NewsID = (SELECT MAX(NewsID) FROM News WHERE NB = @NB);";

                            // 讀取後直接更新 ScanTime
                            using (SqlCommand command = new SqlCommand(sql, connection))
                            {
                                command.Parameters.AddWithValue("@ScanTime", DateTime.Now);
                                command.Parameters.AddWithValue("@NewsImg", ",2024618154829734.jpg");
                                command.Parameters.AddWithValue("@NB", NB);

                                command.ExecuteNonQuery(); // 執行更新
                            }

                            // 查詢最新的資料
                            string Sql = "SELECT NewsStart, NewsEnd, NewsImg, mem_chinese FROM News WHERE NewsID = (SELECT MAX(NewsID) FROM News WHERE NB = @NB);";
                            using (SqlCommand queryCommand = new SqlCommand(Sql, connection))
                            {
                                queryCommand.Parameters.AddWithValue("@NB", NB);

                                using (SqlDataReader reader = queryCommand.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        // 讀取查詢結果
                                        string newsStart = reader["NewsStart"].ToString();
                                        string newsEnd = reader["NewsEnd"].ToString();
                                        string newsImg = reader["NewsImg"].ToString();
                                        string mem_chinese = reader["mem_chinese"].ToString();

                                        // 构建 body 的 JObject
                                        JObject body = new JObject(
                                            new JProperty("type", "box"),
                                            new JProperty("layout", "vertical"),
                                            new JProperty("contents", new JArray(
                                                new JObject(
                                                    new JProperty("type", "text"),
                                                    new JProperty("text", $"日期：{newsStart}~{newsEnd}"),
                                                    new JProperty("color", "#040420"),
                                                    new JProperty("weight", "bold"),
                                                    new JProperty("size", "md")
                                                ),
                                                new JObject(
                                                    new JProperty("type", "text"),
                                                    new JProperty("text", $"設備：{NB}"),
                                                    new JProperty("color", "#040420"),
                                                    new JProperty("weight", "bold"),
                                                    new JProperty("size", "md")
                                                ),
                                                new JObject(
                                                    new JProperty("type", "text"),
                                                    new JProperty("text", $"姓名：{mem_chinese}"),
                                                    new JProperty("color", "#040420"),
                                                    new JProperty("weight", "bold"),
                                                    new JProperty("size", "md")
                                                ),
                                                new JObject(
                                                    new JProperty("type", "text"),
                                                    new JProperty("text", $"檔案：{newsImg.Replace(",", string.Empty)}"),
                                                    new JProperty("color", "#040420"),
                                                    new JProperty("weight", "bold"),
                                                    new JProperty("size", "md")
                                                ),
                                                new JObject(
                                                    new JProperty("type", "text"),
                                                    new JProperty("text", $"狀態：筆電掃描完成"),
                                                    new JProperty("color", "#040420"),
                                                    new JProperty("weight", "bold"),
                                                    new JProperty("size", "md")
                                                )
                                            ))
                                        );

                                        // 构建 output 的 JObject
                                        JObject output = new JObject(
                                            new JProperty("type", "bubble"),
                                            new JProperty("hero", new JObject(
                                                new JProperty("type", "image"),
                                                new JProperty("url", ""),
                                                new JProperty("size", "xl"),
                                                new JProperty("aspectRatio", "50:50")
                                            )),
                                            new JProperty("body", body)
                                        );

                                        // 构建最终的 JObject
                                        JObject contents = new JObject(
                                            new JProperty("type", "flex"),
                                            new JProperty("altText", $"{NB}筆電掃描完成"),
                                            new JProperty("contents", output)
                                        );

                                        // 將結果轉換成 JSON 字串並輸出
                                        string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);

                                        // 回覆消息
                                        ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);
                                    }
                                }
                            }
                        }
                    } 
                }
                if (LineEvent.type == "postback")  //postback
                {
                    string postbackData = LineEvent.postback.data;

                    if (postbackData.Contains("參加"))
                    {
                        // 呼叫 GetUserInfo 方法以獲取使用者名稱和大頭貼
                        //var userInfo = await GetUserProfile(LineEvent.source.userId);

                        // 從返回的元組中獲取使用者名稱
                        //string userName = userInfo.Item1;

                        //取得userInfo
                        //var userInfo = await GetUserProfile(LineEvent.source.userId);
                        ////檢查userInfo
                        //if (userInfo == null)
                        //{
                        //    ReplyMessage(LineEvent.replyToken, "請先加入LineBot好友");
                        //    return NotFound();
                        //}

                        //取得用戶名稱 
                        LineUserInfo userInfo = null;
                        if (LineEvent.source.type.ToLower() == "room")
                            userInfo = isRock.LineBot.Utility.GetRoomMemberProfile(
                                LineEvent.source.roomId, LineEvent.source.userId, ChannelAccessToken);
                        if (LineEvent.source.type.ToLower() == "group")
                            userInfo = isRock.LineBot.Utility.GetGroupMemberProfile(
                                LineEvent.source.groupId, LineEvent.source.userId, ChannelAccessToken);
                        if (LineEvent.source.type.ToLower() == "user")
                            userInfo = isRock.LineBot.Utility.GetUserInfo(
                                 LineEvent.source.userId, ChannelAccessToken);

                        //檢查userInfo
                        if (userInfo == null)
                        {
                            ReplyMessage(LineEvent.replyToken, "請先加入LineBot好友");
                            return NotFound();
                        }

                        // 讀取CSV文件的內容
                        string filePath = HttpContext.Current.Server.MapPath("~/Activity/input.csv");
                        List<string> linesList = new List<string>();

                        // 檢查用戶ID是否已存在
                        bool isUserIdExists = false;
                        List<string> updatedLines = new List<string>();

                        using (StreamReader reader = new StreamReader(filePath, Encoding.GetEncoding("UTF-8")))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                string[] parts = line.Split(',');

                                // 如果找到相同的用戶ID
                                if (parts.Length > 0 && parts[0] == LineEvent.source.userId)
                                {
                                    isUserIdExists = true;

                                    // 更新報名狀態
                                    line = $"{parts[0]},{parts[1]},{postbackData}";
                                }

                                updatedLines.Add(line);
                            }
                        }

                        // 如果使用者 ID 不存在，則將新的行寫入檔案中
                        if (!isUserIdExists)
                        {
                            using (FileStream fs = new FileStream(filePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                            {
                                // 在文件末尾添加新的行，包含用戶ID、姓名和報名狀態
                                using (StreamWriter writer = new StreamWriter(fs, Encoding.GetEncoding("UTF-8")))
                                {
                                    string newLine = $"{LineEvent.source.userId},{HttpUtility.UrlDecode(userInfo.displayName)},{postbackData}";
                                    writer.WriteLine(newLine);
                                }
                            }
                        }
                        else
                        {
                            File.WriteAllLines(filePath, updatedLines, Encoding.GetEncoding("UTF-8"));
                        }


                        // 构建 body 的 JObject
                        JObject body = new JObject(
                            new JProperty("type", "box"),
                            new JProperty("layout", "vertical"),
                            new JProperty("contents", new JArray(
                                new JObject(
                                    new JProperty("type", "text"),
                                    new JProperty("text", $"【{userInfo.displayName}】{postbackData}"),
                                    new JProperty("color", "#040420"),
                                    new JProperty("weight", "bold"),
                                    new JProperty("size", "md")
                                )
                            ))
                        );

                        // 构建 output 的 JObject
                        JObject output = new JObject(
                            new JProperty("type", "bubble"),
                            new JProperty("hero", new JObject(
                                new JProperty("type", "image"),
                                new JProperty("url", ""),
                                new JProperty("size", "xl"),
                                new JProperty("aspectRatio", "50:50")
                            )),
                            new JProperty("body", body)
                        );

                        // 构建最终的 JObject
                        JObject contents = new JObject(
                            new JProperty("type", "flex"),
                            new JProperty("altText", $"【{userInfo.displayName}】{postbackData}"),
                            new JProperty("contents", output)
                        );

                        // 將結果轉換成 JSON 字串並輸出
                        string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);

                        //回覆訊息
                        ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);

                        //ReplyMessage(LineEvent.replyToken, $"【{userInfo.displayName}】{postbackData}");
                    }
                    else//變更已收未收狀態
                    {
                        //管理者
                        List<string> adminNames = new List<string> { "" };
                        //取得userInfo
                        //var userInfo = await GetUserProfile(LineEvent.source.userId);
                        //檢查userInfo
                        //if (userInfo == null)
                        //{
                        //    return NotFound();
                        //}

                        //取得用戶名稱 
                        LineUserInfo userInfo = null;
                        if (LineEvent.source.type.ToLower() == "room")
                            userInfo = isRock.LineBot.Utility.GetRoomMemberProfile(
                                LineEvent.source.roomId, LineEvent.source.userId, ChannelAccessToken);
                        if (LineEvent.source.type.ToLower() == "group")
                            userInfo = isRock.LineBot.Utility.GetGroupMemberProfile(
                                LineEvent.source.groupId, LineEvent.source.userId, ChannelAccessToken);
                        if (LineEvent.source.type.ToLower() == "user")
                            userInfo = isRock.LineBot.Utility.GetUserInfo(
                                 LineEvent.source.userId, ChannelAccessToken);

                        //檢查userInfo
                        if (userInfo == null)
                        {
                            return NotFound();
                        }

                        //檢查管理者名單
                        if (!adminNames.Contains(userInfo.displayName))
                        {
                            return NotFound();
                        }

                        // 讀取CSV文件的內容
                        string filePath = HttpContext.Current.Server.MapPath("~/Order/input.csv");
                        List<string> updatedLines = new List<string>();
                        string status = "未收款";

                        using (StreamReader reader = new StreamReader(filePath, Encoding.GetEncoding("UTF-8")))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                string[] parts = line.Split(',');

                                //if (parts[0]=="總原價") {
                                //    break;
                                //}

                                // 如果找到相同的用戶ID
                                if (parts.Length > 0 && parts[0] == postbackData)
                                {
                                    if (parts[6] == "V")
                                    {
                                        parts[6] = string.Empty;
                                        status = "未收款";
                                    }
                                    else
                                    {
                                        parts[6] = "V";
                                        status = "已收款";
                                    }

                                    line = $"{parts[0]},{parts[1]},{parts[2]},{parts[3]},{parts[4]},{parts[5]},{parts[6]}";
                                }
                                updatedLines.Add(line);
                            }
                        }

                        File.WriteAllLines(filePath, updatedLines, Encoding.GetEncoding("UTF-8"));


                        // 构建 body 的 JObject
                        JObject body = new JObject(
                            new JProperty("type", "box"),
                            new JProperty("layout", "vertical"),
                            new JProperty("contents", new JArray(
                                new JObject(
                                    new JProperty("type", "text"),
                                    new JProperty("text", $"【{postbackData}】{status}"),
                                    new JProperty("color", status == "已收款" ? "#040420" : "#9e0200"),
                                    new JProperty("weight", "bold"),
                                    new JProperty("size", "md")
                                )
                            ))
                        );

                        // 构建 output 的 JObject
                        JObject output = new JObject(
                            new JProperty("type", "bubble"),
                            new JProperty("hero", new JObject(
                                new JProperty("type", "image"),
                                new JProperty("url", status =="已收款" ? "" : ""),
                                new JProperty("size", "xl"),
                                new JProperty("aspectRatio", "50:50")
                            )),
                            new JProperty("body", body)
                        );

                        // 构建最终的 JObject
                        JObject contents = new JObject(
                            new JProperty("type", "flex"),
                            new JProperty("altText", $"【{postbackData}】{status}"),
                            new JProperty("contents", output)
                        );

                        // 將結果轉換成 JSON 字串並輸出
                        string jsonOutput = JsonConvert.SerializeObject(new JArray(contents), Formatting.Indented);


                        //回覆訊息
                        ReplyMessageWithJSON(LineEvent.replyToken, jsonOutput);

                        //ReplyMessage(LineEvent.replyToken, $"【{postbackData}】{status}");
                    }
                }

                return Ok();//response OK
            }
            catch (Exception ex)
            {
                //回覆訊息
                ReplyMessage(AdminUserId, "發生錯誤:\n" + ex.Message);

                //回覆訊息
                //this.PushMessage(AdminUDefaultserId, "發生錯誤:\n" + ex.Message);
                //response OK
                return Ok();
            }
        }

        // 參加&不參加事件(超連結)
        [Route("api/JoinActivity")]
        [HttpGet]
        public IHttpActionResult JoinActivity(string userId, string userName, string isJoin)
        {
            try
            {
                // 讀取CSV文件的內容
                string filePath = HttpContext.Current.Server.MapPath("~/Activity/input.csv");
                List<string> linesList = new List<string>();

                // 檢查用戶ID是否已存在
                bool isUserIdExists = false;
                List<string> updatedLines = new List<string>();

                using (StreamReader reader = new StreamReader(filePath, Encoding.GetEncoding("UTF-8")))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] parts = line.Split(',');

                        // 如果找到相同的用戶ID
                        if (parts.Length > 0 && parts[0] == userId)
                        {
                            isUserIdExists = true;

                            // 更新報名狀態
                            line = $"{parts[0]},{parts[1]},{isJoin}";
                        }

                        updatedLines.Add(line);
                    }
                }

                // 如果使用者 ID 不存在，則將新的行寫入檔案中
                if (!isUserIdExists)
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                    {
                        // 在文件末尾添加新的行，包含用戶ID、姓名和報名狀態
                        using (StreamWriter writer = new StreamWriter(fs, Encoding.GetEncoding("UTF-8")))
                        {
                            string newLine = $"{userId},{HttpUtility.UrlDecode(userName)},{isJoin}";
                            writer.WriteLine(newLine);
                        }
                    }
                }
                else
                {
                    File.WriteAllLines(filePath, updatedLines, Encoding.GetEncoding("UTF-8"));
                }

                // 回覆訊息
                //var bot = new Bot(channelAccessToken);
                //bot.ReplyMessage(userId, $"【{HttpUtility.UrlDecode(userName)}】【{isJoin}】登記完成");

                // 返回成功頁面或其他操作
                return Ok($"【{HttpUtility.UrlDecode(userName)}】【{isJoin}】登記完成");
            }
            catch (Exception ex)
            {
                // 返回錯誤訊息
                return InternalServerError(ex);
            }
        }

        /// <summary>
        /// 使用者資訊
        /// </summary>
        /// <param name="userId"></param>
        /// <returns></returns>
        public async Task<(string, string)> GetUserInfo(string userId)
        {
            // LINE 使用者資訊 API 網址
            string userInfoUrl = $"https://api.line.me/v2/bot/profile/{userId}";

            // 使用 Channel Access Token 進行 API 請求
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {this.channelAccessToken}");

                // 發送 GET 請求
                HttpResponseMessage response = await client.GetAsync(userInfoUrl);

                // 確認回應成功
                if (response.IsSuccessStatusCode)
                {
                    // 讀取回應內容
                    string responseBody = await response.Content.ReadAsStringAsync();

                    // 解析 JSON 回應
                    JObject jsonResponse = JObject.Parse(responseBody);

                    // 取得使用者名稱和大頭貼 URL
                    string userName = jsonResponse["displayName"].ToString();
                    string profilePictureUrl = jsonResponse["pictureUrl"].ToString();

                    return (userName, profilePictureUrl);
                }
                else
                {
                    // 若回應不成功，返回空字串或處理錯誤
                    return ("", "");
                }
            }
        }

        HttpClient httpClient = new HttpClient();
        private async Task<UserProfile> GetUserProfile(string userId)
        {
            //GET https://api.line.me/v2/bot/profile/{userId}
            var httpRequestMessage = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri($"https://api.line.me/v2/bot/profile/{userId}"),
                Headers = {
                    { "Authorization", $"Bearer {ChannelAccessToken}" },
                }
            };
            var result = await httpClient.SendAsync(httpRequestMessage);

            var content = await result.Content.ReadAsStringAsync();

            var profile = JsonConvert.DeserializeObject<UserProfile>(content);

            return profile;
        }

        /// <summary>
        /// 不在勤名單
        /// </summary>
        /// <param name="workDate"></param>
        /// <returns></returns>
        private static string GetVacationListFromDatabase(string workDate)
        {
            var vacationList = new List<string>();

            try
            {
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["fastDbContext"].ConnectionString))
                {
                    connection.Open();
                        
                    string query = $"SELECT emp_sno, depname, osname, empname, title, work_date, stime1, etime1, stime2, etime2, leave, leave2, agent1, ag1_depname, ag1_empname, ag1_title, agent2, ag2_depname, ag2_empname, ag2_title, leave1_app_sno, leave2_app_sno, leave3, leave4, leave3_app_sno, leave4_app_sno,COALESCE((SELECT field_data FROM app_data WHERE serial='8' AND app_sno = w.leave1_app_sno),(SELECT field_data FROM app_data WHERE serial='8' AND app_sno = w.leave2_app_sno),(SELECT field_data FROM app_data WHERE serial='8' AND app_sno = w.leave3_app_sno),(SELECT field_data FROM app_data WHERE serial='8' AND app_sno = w.leave4_app_sno)) AS field_data FROM work_time_view w LEFT JOIN app_data a ON w.leave1_app_sno = a.app_sno OR w.leave2_app_sno = a.app_sno OR w.leave3_app_sno = a.app_sno OR w.leave4_app_sno = a.app_sno WHERE work_date = @WorkDate AND (leave IS NOT NULL OR leave2 IS NOT NULL OR leave3 IS NOT NULL OR leave4 IS NOT NULL) AND a.serial='4'";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        //查詢日期
                        command.Parameters.AddWithValue("@WorkDate", workDate);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            int index = 1;
                            while (reader.Read())
                            {
                                string employeeName = ReplaceMiddleCharacter(reader["empname"].ToString().Trim(), 'O');
                                string depname = reader["depname"].ToString().Trim();
                                string workDateResult = reader["work_date"].ToString().Trim();
                                string stime1 = reader["stime1"].ToString().Trim();
                                string etime1 = reader["etime1"].ToString().Trim();
                                string stime2 = reader["stime2"].ToString().Trim();
                                string etime2 = reader["etime2"].ToString().Trim();
                                string leave = reader["leave"].ToString().Trim();
                                string leave2 = reader["leave2"].ToString().Trim();
                                string leave4 = reader["leave4"].ToString().Trim();
                                string field_data = string.IsNullOrEmpty(reader["field_data"].ToString().Trim()) ? string.Empty : $"({reader["field_data"].ToString().Trim()})";

                                string leave1_app_sno = reader["leave1_app_sno"].ToString().Trim();
                                string leaveType = "";

                                if (!string.IsNullOrEmpty(leave) && !string.IsNullOrEmpty(leave2))
                                {
                                    if (leave == leave2)
                                    {
                                        leaveType = $"全日:{leave}";
                                    }
                                    else
                                    {
                                        leaveType = $"上午:{leave} 下午:{leave2}";
                                    }
                                    
                                }
                                else if (string.IsNullOrEmpty(leave) && !string.IsNullOrEmpty(leave2))
                                {
                                    leaveType = $"下午:{leave2}";
                                }
                                else if (!string.IsNullOrEmpty(leave) && string.IsNullOrEmpty(leave2))
                                {
                                    leaveType = $"上午:{leave}";
                                }
                                else if (string.IsNullOrEmpty(leave) && string.IsNullOrEmpty(leave2))
                                {
                                    leaveType = field_data;
                                    leaveType = leaveType + $"({leave4}分)";
                                }

                                // 将字符串每17个字符加入换行符
                                int length = 17;
                                StringBuilder sb = new StringBuilder();
                                for (int i = 0; i < field_data.Length; i += length)
                                {
                                    if (i + length < field_data.Length)
                                    {
                                        sb.Append(field_data.Substring(i, length)).Append(Environment.NewLine);
                                    }
                                    else
                                    {
                                        sb.Append(field_data.Substring(i));
                                    }
                                }

                                field_data = sb.ToString();
                                vacationList.Add($"{index}.{depname} {employeeName} {leaveType}{"\n"}{field_data}{"\n"}");

                                index++;
                            }
                        }
                    }
                }
                return string.Join("", vacationList);
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 加班名單
        /// </summary>
        /// <param name="workDate"></param>
        /// <returns></returns>
        private static string GetOverTimeListFromDatabase(string workDate)
        {
            var vacationList = new List<string>();

            string query1 = $"SELECT * FROM app_data WHERE app_sno IN (SELECT sno FROM application WHERE flow_sno = '7') AND (serial = '2' AND ISDATE(field_data) = 1 AND CONVERT(date, CONVERT(datetime, field_data)) = '{workDate}') ORDER BY app_sno DESC, serial";
            DataTable dt2 = new DataTable();
            dt2.Columns.Add("empname", typeof(string));
            dt2.Columns.Add("sdate", typeof(string));
            dt2.Columns.Add("edate", typeof(string));
            dt2.Columns.Add("time", typeof(string));
            dt2.Columns.Add("remark", typeof(string));
            dt2.Columns.Add("ny", typeof(string));

            foreach (DataRow row in GetDataTable(query1).Rows)
            {
                string app_sno = row["app_sno"].ToString();

                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["fastDbContext"].ConnectionString))
                {
                    connection.Open();

                    string query2 = $"SELECT app_data.sno,app_data.app_sno,CASE WHEN app_data.serial = '1' THEN employees.empname ELSE app_data.field_data END AS field_data,app_data.serial FROM app_data LEFT JOIN employees ON CONVERT(VARCHAR, app_data.field_data) = CONVERT(VARCHAR, employees.sno) WHERE app_data.app_sno = '{app_sno}'";

                    using (SqlCommand command = new SqlCommand(query2, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            int rowCount = 0;
                            string empname = string.Empty, sdate = string.Empty, edate = string.Empty, remark = string.Empty, ny = string.Empty, time = string.Empty;

                            while (reader.Read())
                            {
                                rowCount++;

                                if (rowCount == 1)
                                    empname = reader["field_data"].ToString().Trim();
                                if (rowCount == 2)
                                {
                                    DateTime dateTime = DateTime.Parse(reader["field_data"].ToString().Trim());
                                    sdate = dateTime.ToString("HH:mm");
                                }  
                                if (rowCount == 3)
                                {
                                    DateTime dateTime = DateTime.Parse(reader["field_data"].ToString().Trim());
                                    edate = dateTime.ToString("HH:mm");
                                }
                                if (rowCount == 4)
                                    time = reader["field_data"].ToString().Trim();
                                if (rowCount == 5)
                                    remark = reader["field_data"].ToString().Trim();
                                if (rowCount == 11)
                                    ny = reader["field_data"].ToString().Trim().Equals("是") ? "*申請補休" : "";
                                if (rowCount == 12)
                                    dt2.Rows.Add(empname, sdate, edate, time, remark, ny);
                            }
                        }
                    }
                }
            }

            int index = 1;
            foreach (DataRow row in dt2.Rows)
            {
                vacationList.Add($"{index}.{ReplaceMiddleCharacter(row["empname"].ToString(), 'O')} ({row["sdate"].ToString()}~{row["edate"].ToString()})+{row["time"].ToString()}hr{"\n"}{row["remark"].ToString()}{"\n"}{row["ny"].ToString()}{"\n"}");
                index++;
            }
            return string.Join("", vacationList);
        }

        static string ReplaceMiddleCharacter(string input, char replacement)
        {
            if (input.Length % 2 == 0 || input.Length < 3)
            {
                return input;
            }

            int middleIndex = input.Length / 2;
            string modifiedString = input.Substring(0, middleIndex) + replacement + input.Substring(middleIndex + 1);
            return modifiedString;
        }

        /// <summary>
        /// 取得星期
        /// </summary>
        /// <param name="dayOfWeek"></param>
        /// <returns></returns>
        static string ConvertToChineseDayOfWeek(DayOfWeek dayOfWeek)
        {
            switch (dayOfWeek)
            {
                case DayOfWeek.Sunday:
                    return "日";
                case DayOfWeek.Monday:
                    return "一";
                case DayOfWeek.Tuesday:
                    return "二";
                case DayOfWeek.Wednesday:
                    return "三";
                case DayOfWeek.Thursday:
                    return "四";
                case DayOfWeek.Friday:
                    return "五";
                case DayOfWeek.Saturday:
                    return "六";
                default:
                    return ""; // 異常情況
            }
        }

        public static DataTable GetDataTable(string sql)
        {
            // 创建 DataTable
            DataTable dataTable = new DataTable();

            // 创建连接对象和适配器
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["fastDbContext"].ConnectionString))
            {
                // 打开连接
                connection.Open();

                // 创建适配器
                using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connection))
                {
                    // 使用适配器填充 DataTable
                    adapter.Fill(dataTable);
                }
            }

            // 返回填充后的 DataTable
            return dataTable;
        }
    }

    public class UserProfile
    {
        public string displayName { get; set; }
        public string userId { get; set; }
        public string pictureUrl { get; set; }
        public string statusMessage { get; set; }
    }
}