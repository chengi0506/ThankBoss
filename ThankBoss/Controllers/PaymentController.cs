using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ThankBoss.Controllers
{
    public class PaymentController : Controller
    {
        // GET: Payment
        public ActionResult Index(string userId)
        {
            // 在這裡處理使用者 ID，例如將其存儲到資料庫中或進行其他操作
            ViewBag.UserId = userId;

            // 返回付款頁面視圖
            return View();
        }
    }
}