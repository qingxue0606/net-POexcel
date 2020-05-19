using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Logging;
using POexcel.Models;

namespace POexcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private String connString;
        private readonly IWebHostEnvironment _webHostEnvironment;


        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment webHostEnvironment)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "AppData\\" + "demo_poexcel.db";
            connString = "Data Source=" + dataPath;

        }

        [Route("/Login")]
        public IActionResult Login()
        {
            String tz = Request.Query["tz"];
            if (tz != null && tz.Length > 0)
            {
                if ((Request.Form["TextUserName"].ToString() == "admin") && (Request.Form["TextPassword"].ToString() == "123"))
                {
                    HttpContext.Session.SetString("UserName", "admin");//放置string数据
                    return Redirect("/");
                }
            }
            return View();
        }
        [Route("/Logout")]
        public IActionResult Logout()
        {
            HttpContext.Session.SetString("UserName", "");
            return Redirect("/");
        }

        public IActionResult Index()
        {
            //获取index.aspx页面传递过来参数的值
            String userName = HttpContext.Session.GetString("UserName");
            if (userName == null || userName.Length <= 0)
            {
                return Redirect("/Login");
            }

            string sql = "select * from OrderMaster order by id desc ";
            SqliteConnection conn = new SqliteConnection(connString);

            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            StringBuilder strHtmls = new StringBuilder();


            while (dr.Read())
            {
                strHtmls.Append("<tr>\n");
                strHtmls.Append("<td>" + dr["OrderNum"] + "</td>");
                if (dr["OrderDate"] != null && dr["OrderDate"].ToString().Trim().Length > 0)
                {
                    strHtmls.Append("<td>" + DateTime.Parse(dr["OrderDate"].ToString()).ToShortDateString() + "</td>\n");
                }
                else
                {
                    strHtmls.Append("<td>&nbsp;</td>\n");

                }
                strHtmls.Append("<td>" + dr["CustName"] + "</td>\n");
                strHtmls.Append("<td>" + dr["SalesName"] + "</td>\n");
                if (dr["Amount"] != null && dr["Amount"].ToString().Trim().Length > 0)
                {
                    strHtmls.Append(" <td style='text-align:right;padding-right:5px;'>" + string.Format("{0:C}", dr["Amount"]) + "</td>\n");
                }
                else
                {
                    strHtmls.Append(" <td>&nbsp;</td>\n");
                }
                strHtmls.Append("<td>\n");
                strHtmls.Append("<div class='ul-page'>\n");
                strHtmls.Append("<a href=\"javascript:POBrowser.openWindowModeless('Order/OpenOrder?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\" >修改</a>|<a href= \"javascript:POBrowser.openWindowModeless('Order/ViewOrder?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\" >只读查看^打印</a>|<a  onclick='Delete(" + dr["ID"] + ")' >删除</a>\n");
                strHtmls.Append("</div>\n");
                strHtmls.Append("</td>\n");
                strHtmls.Append("</tr>\n");
            }

            dr.Close();
            conn.Close();
            ViewBag.strHtmls = strHtmls;
            ViewBag.Data = DateTime.Now.ToShortDateString() + "    " + "星期" + DateTime.Now.DayOfWeek.ToString(("D"));
            ViewBag.userName = HttpContext.Session.GetString("UserName");
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
