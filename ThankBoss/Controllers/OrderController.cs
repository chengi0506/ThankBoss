using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ThankBoss.Models;

namespace ThankBoss.Controllers
{
    public class OrderController : Controller
    {
        // GET: Order
        public ActionResult Index()
        {
            return View(new OrderModel());
        }

        [HttpPost]
        public ActionResult GenerateJson(OrderModel model)
        {
            if (!ModelState.IsValid)
            {
                return View("Index", model);
            }

            // 解析輸入內容並創建相應的 JSON 物件
            JArray sectionsArray;
            try
            {
                sectionsArray = JArray.Parse(model.InputContent);
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(nameof(model.InputContent), "Invalid JSON format: " + ex.Message);
                return View("Index", model);
            }

            // 將 JSON 物件序列化為字符串
            string jsonOutput = sectionsArray.ToString();

            // 將生成的 JSON 字符串存儲到 ViewBag 中，以便在視圖中顯示
            ViewBag.JsonOutput = jsonOutput;

            return View("Index", model);
        }
    }
}