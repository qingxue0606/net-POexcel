using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;
using Microsoft.VisualBasic;
using PageOfficeNetCore.ExcelReader;

namespace POexcel.Controllers.Order
{
    public class OrderController : Controller
    {

        protected string strErrHtml = "";//错误提示
        private string connString;

        private readonly IWebHostEnvironment _webHostEnvironment;

        public OrderController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string dataPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            dataPath = dataPath.Substring(0, dataPath.Length - 7) + "appData\\" + "demo_poexcel.db";
            connString = "Data Source=" + dataPath;
        }
        public IActionResult OpenOrder()
        {

            string id = Request.Query["ID"];
            string sql = "select * from OrderMaster where ID=" + id;
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();

            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            // 填充数据
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //workBook.DisableSheetDoubleClick = true;
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("销售订单");
            if (dr.Read()) {

                sheet.OpenCell("D5").Value = dr["CustName"].ToString();
                sheet.OpenCell("D5").SubmitName = "CustName";//单元格提交数据 
                sheet.OpenCell("I5").Value = dr["OrderNum"].ToString();
                sheet.OpenCell("I5").SubmitName = "OrderNum";//单元格提交数据
                sheet.OpenCell("D6").Value = dr["CustDistrict"].ToString();
                sheet.OpenCell("D6").SubmitName = "CustDistrict";//单元格提交数据
                sheet.OpenCell("I6").Value = dr["OrderDate"].ToString();
                sheet.OpenCell("I6").SubmitName = "OrderDate";//单元格提交数据
                sheet.OpenCell("D18").Value = dr["MakerName"].ToString();
                sheet.OpenCell("D18").SubmitName = "UserName";//单元格提交数据
                sheet.OpenCell("H18").Value = dr["SalesName"].ToString();
                sheet.OpenCell("H18").SubmitName = "SalesName";//单元格提交数据
                sheet.OpenCell("I16").SubmitName = "Amount";//单元格提交数据
                sheet.OpenCell("I16").ReadOnly = true;//将Excel模版中有公式的单元格设置为只读格式，以免覆盖掉公式

                string sql2 = "select * from OrderDetail where OrderID =" + dr["ID"];

                SqliteCommand cmd2 = new SqliteCommand(sql2, conn);
                cmd2.ExecuteNonQuery();
                SqliteDataReader dr2 = cmd2.ExecuteReader();


                //定义table对象
                PageOfficeNetCore.ExcelWriter.Table tableD = sheet.OpenTable("D9:D15");//定义table对象
                tableD.ReadOnly = true;//将table设置成只读

                PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTable("C9:H15");
                table.SubmitName = "OrderDetail"; //表提交数据

                String proCode = "";
                String type = "";
                String unit = "";
                String quantity = "";
                String price = "";

                while (dr2.Read()) {

                    proCode = dr2["ProductCode"].ToString();
                    type = dr2["ProductType"].ToString();
                    unit = dr2["Unit"].ToString();
                    quantity = dr2["Quantity"].ToString();
                    price = dr2["Price"].ToString();
                    table.DataFields[0].Value = proCode;
                    table.DataFields[2].Value = type;
                    table.DataFields[3].Value = type;
                    table.DataFields[4].Value = quantity;
                    table.DataFields[5].Value = price;
                    table.NextRow();
                }

                dr2.Close();
                table.Close();//关闭table

            }

            dr.Close();
            conn.Close();

            pageofficeCtrl.SetWriter(workBook);

            pageofficeCtrl.AddCustomToolButton("保存并关闭", "Store", 1);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "SetScreen", 4);
            pageofficeCtrl.BorderStyle = PageOfficeNetCore.BorderStyleType.BorderThin;
            string fileName = "OrderForm.xls";
            //设置保存页面
            pageofficeCtrl.SaveDataPage = "UpdateOrder?ID=" + id;
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.xlsSubmitForm, "admin");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }

        public IActionResult NewOrder()
        {

            
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";


            // 填充数据
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //workBook.DisableSheetDoubleClick = true;
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("销售订单");
           

            sheet.OpenCell("D5").SubmitName = "CustName";
            sheet.OpenCell("I5").SubmitName = "OrderNum";
            sheet.OpenCell("D6").SubmitName = "CustDistrict";
            sheet.OpenCell("I6").SubmitName = "OrderDate";
            sheet.OpenCell("I6").Value = DateTime.Now.ToShortDateString();
            sheet.OpenCell("D18").Value = Convert.ToString("admin");
            sheet.OpenCell("D18").SubmitName = "UserName";
            sheet.OpenCell("H18").SubmitName = "SalesName";
     
            sheet.OpenTable("C9:H15").SubmitName = "OrderDetail";
            sheet.OpenCell("I16").SubmitName = "Amount";
         

            sheet.OpenCell("I6").ReadOnly = true;//将Excel模版中有公式的单元格设置为只读格式，以免覆盖掉公式

            pageofficeCtrl.SetWriter(workBook);

            pageofficeCtrl.AddCustomToolButton("保存并关闭", "Store", 1);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "SetScreen", 4);
            pageofficeCtrl.BorderStyle = PageOfficeNetCore.BorderStyleType.BorderThin;
            string fileName = "OrderForm.xls";
            //设置保存页面
            pageofficeCtrl.SaveDataPage = "UpdateOrder?";
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.xlsSubmitForm, "admin");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            return View();
        }

























        public IActionResult ViewOrder()
        {

            string id = Request.Query["ID"];
            string sql = "select * from OrderMaster where ID=" + id;
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();

            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            // 填充数据
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            //workBook.DisableSheetDoubleClick = true;
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("销售订单");
            if (dr.Read())
            {

                sheet.OpenCell("D5").Value = dr["CustName"].ToString();
                sheet.OpenCell("D5").SubmitName = "CustName";//单元格提交数据 
                sheet.OpenCell("I5").Value = dr["OrderNum"].ToString();
                sheet.OpenCell("I5").SubmitName = "OrderNum";//单元格提交数据
                sheet.OpenCell("D6").Value = dr["CustDistrict"].ToString();
                sheet.OpenCell("D6").SubmitName = "CustDistrict";//单元格提交数据
                sheet.OpenCell("I6").Value = dr["OrderDate"].ToString();
                sheet.OpenCell("I6").SubmitName = "OrderDate";//单元格提交数据
                sheet.OpenCell("D18").Value = dr["MakerName"].ToString();
                sheet.OpenCell("D18").SubmitName = "UserName";//单元格提交数据
                sheet.OpenCell("H18").Value = dr["SalesName"].ToString();
                sheet.OpenCell("H18").SubmitName = "SalesName";//单元格提交数据
                sheet.OpenCell("I16").SubmitName = "Amount";//单元格提交数据
                sheet.OpenCell("I16").ReadOnly = true;//将Excel模版中有公式的单元格设置为只读格式，以免覆盖掉公式

                string sql2 = "select * from OrderDetail where OrderID =" + dr["ID"];

                SqliteCommand cmd2 = new SqliteCommand(sql2, conn);
                cmd2.ExecuteNonQuery();
                SqliteDataReader dr2 = cmd2.ExecuteReader();


                //定义table对象
                PageOfficeNetCore.ExcelWriter.Table tableD = sheet.OpenTable("D9:D15");//定义table对象
                tableD.ReadOnly = true;//将table设置成只读

                PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTable("C9:H15");
                table.SubmitName = "OrderDetail"; //表提交数据

                String proCode = "";
                String type = "";
                String unit = "";
                String quantity = "";
                String price = "";

                while (dr2.Read())
                {

                    proCode = dr2["ProductCode"].ToString();
                    type = dr2["ProductType"].ToString();
                    unit = dr2["Unit"].ToString();
                    quantity = dr2["Quantity"].ToString();
                    price = dr2["Price"].ToString();
                    table.DataFields[0].Value = proCode;
                    table.DataFields[2].Value = type;
                    table.DataFields[3].Value = type;
                    table.DataFields[4].Value = quantity;
                    table.DataFields[5].Value = price;
                    table.NextRow();
                }

                dr2.Close();
                table.Close();//关闭table

            }
            dr.Close();
            conn.Close();
            workBook.DisableSheetSelection = true;
            pageofficeCtrl.SetWriter(workBook);

            //添加自定义菜单
            pageofficeCtrl.AddCustomToolButton("打印", "Print", 6);
            pageofficeCtrl.AddCustomToolButton("打印预览", "PrintPreView", 7);
            pageofficeCtrl.AddCustomToolButton("页面设置", "SetPage", 3);
            pageofficeCtrl.AddCustomMenuItem("|", "", true);
            pageofficeCtrl.AddCustomToolButton("另存到本机", "StoreAs", 1);
            pageofficeCtrl.AddCustomMenuItem("|", "", true);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "SetScreen", 4);
            string fileName = "OrderForm.xls";
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "SaveDoc?ID=" + id;
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.xlsReadOnly, "admin");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }




        public IActionResult OrderStat()
        {

            string sql = "SELECT OrderMaster.SalesName as SalesName , OrderDetail.ProductName as ProductName , sum(OrderDetail.Quantity) as Quantity, sum(OrderDetail.Price * OrderDetail.Quantity) as amount "
            + "from OrderMaster,OrderDetail "
            + " where OrderMaster.ID = OrderDetail.OrderID  and OrderMaster.SalesName in('阿土伯','金贝贝','钱夫人','孙小美')  "
            + " group by OrderMaster.SalesName, OrderDetail.ProductName";;
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();

            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            // 填充数据
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("统计图表");
            String salesName = "";
            String preSalesName = "";
            String qua = "";
            String proName = "";
            String amount = "";
            int columnId = 5;
            int n = 0;
            while (dr.Read())
            {

                salesName = dr["SalesName"].ToString();
                qua = dr["Quantity"].ToString();
                proName = dr["ProductName"].ToString(); 
                amount = dr["amount"].ToString();


                if (!salesName.Equals(preSalesName))
                {
                    columnId = 5 + n * 4;
                    n++;
                }
                preSalesName = salesName;
                sheet.OpenCell("B" + columnId).Value=salesName;

                if (proName.Equals("笔记本"))
                {
                    sheet.OpenCell("C" + columnId).Value=proName;
                    sheet.OpenCell("D" + columnId).Value=qua;
                    sheet.OpenCell("E" + columnId).Value=amount;
                }

                if (proName.Equals("服务器"))
                {
                    sheet.OpenCell("C" + (columnId + 1)).Value=proName;
                    sheet.OpenCell("D" + (columnId + 1)).Value=qua;
                    sheet.OpenCell("E" + (columnId + 1)).Value=amount;
                }
                if (proName.Equals("路由器"))
                {
                    sheet.OpenCell("C" + (columnId + 2)).Value=proName;
                    sheet.OpenCell("D" + (columnId + 2)).Value=qua;
                    sheet.OpenCell("E" + (columnId + 2)).Value=amount;
                }

            }
            dr.Close();
            conn.Close();

            workBook.DisableSheetSelection = true;
            pageofficeCtrl.SetWriter(workBook);

            //添加自定义菜单
            pageofficeCtrl.AddCustomToolButton("打印", "Print", 6);
            pageofficeCtrl.AddCustomToolButton("打印预览", "PrintPreView", 7);
            pageofficeCtrl.AddCustomToolButton("页面设置", "SetPage", 3);
            pageofficeCtrl.AddCustomMenuItem("|", "", true);
            pageofficeCtrl.AddCustomToolButton("另存到本机", "StoreAs", 1);
            pageofficeCtrl.AddCustomMenuItem("|", "", true);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "SetScreen", 4);
            string fileName = "OrderReport.xls";
            pageofficeCtrl.Caption = "统计图表";
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.xlsSubmitForm, "admin");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }



        public IActionResult OrderStat2()
        {

            string sql = "SELECT OrderNum,OrderDate,CustName,SalesName,Amount from OrderMaster order by ID desc"; 
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();

            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            // 填充数据
            PageOfficeNetCore.ExcelWriter.Workbook workBook = new PageOfficeNetCore.ExcelWriter.Workbook();
            PageOfficeNetCore.ExcelWriter.Sheet sheet = workBook.OpenSheet("查询表");
            int rowCount = 0;//记录行数
            String salesName = "";
            String date = "";
            String orderNum = "";
            String custName = "";
            String amount = "";
            double totalMoney = 0.00;
            


            while (dr.Read())
            {

                orderNum = dr["OrderNum"].ToString();
                date = dr["OrderDate"].ToString(); 
                custName=dr["CustName"].ToString(); 
                salesName = dr["SalesName"].ToString(); 
                amount = dr["Amount"].ToString(); 

                sheet.OpenCell("B" + (5 + rowCount)).Value=orderNum;

                sheet.OpenCell("C" + (5 + rowCount)).Value = date;
               
                sheet.OpenCell("D" + (5 + rowCount)).Value=custName;
                sheet.OpenCell("E" + (5 + rowCount)).Value=salesName;
                sheet.OpenCell("F" + (5 + rowCount)).Value=amount;



         
                totalMoney += double.Parse(amount);


                if (rowCount % 2 == 0)
                {
                    //设置背景色
                    sheet.OpenTable(
                            "B" + (5 + rowCount) + ":F" + (5 + rowCount))
                            .BackColor= Color.White;
                }
                rowCount++;

            }
            dr.Close();
            conn.Close();


            //设置前景色
            sheet.OpenTable("B5:F" + (rowCount + 4)).BackColor = Color.Gray;



            //水平方向对齐方式
            sheet.OpenTable("B5:F" + (rowCount + 4)).HorizontalAlignment=
                    PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignLeft;
            sheet.OpenTable("C5:C" + (rowCount + 4)).HorizontalAlignment=
                     PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignCenter;
            sheet.OpenTable("E5:E" + (rowCount + 4)).HorizontalAlignment=
                     PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignCenter;
            sheet.OpenTable("F5:F" + (rowCount + 4)).HorizontalAlignment=
                     PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignRight;
            //竖直方向对齐方式
            sheet.OpenTable("B5:F" + (rowCount + 4)).VerticalAlignment=
                     PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;

            //合计：

            //合并单元格
            sheet.OpenTable("B" + (rowCount + 5) + ":F" + (rowCount + 5))
                    .Merge();
            //行高
            sheet.OpenTable("B5:F" + (rowCount + 6)).RowHeight=18;
            sheet.OpenTable("B" + (rowCount + 5) + ":F" + (rowCount + 6))
                    .HorizontalAlignment= PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignLeft;
            sheet.OpenTable("B" + (rowCount + 5) + ":F" + (rowCount + 6)).VerticalAlignment= PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;

            sheet.OpenCell("B" + (rowCount + 6)).Value="合计";
            sheet.OpenTable("C" + (rowCount + 6) + ":E" + (rowCount + 6))
                    .Merge();

            sheet.OpenCell("F" + (rowCount + 6)).Value= totalMoney.ToString();



            sheet.OpenTable("F" + (rowCount + 6) + ":F" + (rowCount + 6))
                    .HorizontalAlignment= PageOfficeNetCore.ExcelWriter.XlHAlign.xlHAlignRight;
            sheet.OpenTable("B" + (rowCount + 6) + ":F" + (rowCount + 6))
                    .VerticalAlignment= PageOfficeNetCore.ExcelWriter.XlVAlign.xlVAlignCenter;

            //设置字体：大小、名称
            sheet.OpenTable("B5:F" + (rowCount + 6)).Font.Size=9;
            sheet.OpenTable("B5:F" + (rowCount + 6)).Font.Name="宋体";

            //设置Table的边框样式：样式、宽度、颜色(多种边框样式重叠时，需创建Table对象才可实现样式的叠加覆盖)
            PageOfficeNetCore.ExcelWriter.Table table = sheet.OpenTable("B" + (rowCount + 6) + ":F"+ (rowCount + 6));
            table.Border.BorderType= PageOfficeNetCore.ExcelWriter.XlBorderType.xlTopEdge;
            table.Border.Weight= PageOfficeNetCore.ExcelWriter.XlBorderWeight.xlThin;
            table.Border.LineColor=Color.Red;

            table.Close();



            workBook.DisableSheetSelection = true;
            pageofficeCtrl.SetWriter(workBook);

            //添加自定义菜单
            pageofficeCtrl.AddCustomToolButton("打印", "Print", 6);
            pageofficeCtrl.AddCustomToolButton("打印预览", "PrintPreView", 7);
            pageofficeCtrl.AddCustomToolButton("页面设置", "SetPage", 3);
            pageofficeCtrl.AddCustomMenuItem("|", "", true);
            pageofficeCtrl.AddCustomToolButton("另存到本机", "StoreAs", 1);
            pageofficeCtrl.AddCustomMenuItem("|", "", true);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "SetScreen", 4);
            string fileName = "OrderQuery.xls";
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.xlsSubmitForm, "admin");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }





        public async Task<ActionResult> UpdateOrder()
        {
            
            PageOfficeNetCore.ExcelReader.Workbook workBook = new PageOfficeNetCore.ExcelReader.Workbook(Request, Response);
            await workBook.LoadAsync();
            string id = Request.Query["ID"];

            string sql = "select * from OrderMaster where ID=" + id;
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            
            PageOfficeNetCore.ExcelReader.Sheet sheet = workBook.OpenSheet("销售订单");


            if (id != null && id.Length > 0) {
                int num;
                //保存客户信息
                num = UpdateOrInsertCustInfo(cmd, id, workBook, sheet, 0);
                if (num > 0)//保存成功
                {
                    int resDelete = 0;//要删除的记录条数

                    //删除当前orderID下的产品数据
                    sql = "delete from OrderDetail where OrderId = " + id;
                    try
                    {
                        cmd.CommandText = sql;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        strErrHtml += "删除客户ID为" + id + "的产品订单信息失败，失败原因为：" + ex.Message + "\n";
                        resDelete = -1;
                    }

                    //删除成功或无数据可删除时
                    if (resDelete >= 0)
                    {
                        //插入产品信息
                        InsertProductInfo(cmd, workBook, sheet, id);
                    }
                }
                else
                {
                    strErrHtml += "<br>客户信息保存失败！";
                }
            }

            else
            {
            
                int maxId = 0;//OrderMaster表中最大ID号
                sql = "select max(ID) from OrderMaster ";
                cmd.CommandText = sql;




                try
                {
                    object obj = cmd.ExecuteScalar();
                    if (obj != null)
                    {
                        maxId = int.Parse(obj.ToString());
                        //保存客户信息
                        if (UpdateOrInsertCustInfo(cmd, "", workBook, sheet, maxId) > 0)
                        {
                            //插入产品信息
                            InsertProductInfo(cmd, workBook, sheet, (maxId + 1).ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    strErrHtml += "新建订单失败，失败原因为：" + ex.Message;
                }
               
            }


            if (strErrHtml.Length > 0)
            {
                strErrHtml = "\n" + strErrHtml;
                workBook.ShowPage(410, 260);
                workBook.CustomSaveResult = "error";
                await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(strErrHtml));

            }
            workBook.Close();
            
            conn.Close();

            return Content("OK");
        }

        private void InsertProductInfo(SqliteCommand cmd, Workbook workBook, Sheet sheet, string orderId)
        {
            PageOfficeNetCore.ExcelReader.Table table = sheet.OpenTable("OrderDetail");
            while (!table.EOF)
            {
                //根据当前OrderID重新插入产品数据
                string sql = "insert into OrderDetail(OrderID, ProductCode, ProductName, ProductType, Unit, Quantity, Price) values(" + orderId;
                if (!table.DataFields.IsEmpty)//数据字段非空时
                {
                    int qua = 0;//数量
                    if (table.DataFields[4].Value.Trim().Length > 0 && int.TryParse(table.DataFields[4].Value.Trim(), out qua))
                    {
                        qua = int.Parse(table.DataFields[4].Value.Trim());
                    }
                    float price = 0.00f;//单价
                    if (float.TryParse(table.DataFields[5].Value.Trim(), out price))
                    {
                        price = float.Parse(table.DataFields[5].Value.Trim());
                    }
                    sql += ",'" + table.DataFields[0].Value + "','" + table.DataFields[1].Value + "','" +
                        table.DataFields[2].Value.Trim() + "','" + table.DataFields[3].Value.Trim() + "'," +
                        qua + ",'" + price + "')";
                    try
                    {
                        cmd.CommandText = sql;
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        strErrHtml += "客户ID为" + orderId + "的产品订单信息添加失败，失败原因为：" + ex.Message + "\n"; ;
                    }
                }

                table.NextRow();//跳到下一行
            }
            table.Close();//关闭table
        }

        private int UpdateOrInsertCustInfo(SqliteCommand cmd, string cid, Workbook workBook, Sheet sheet, int maxId)
            {

                string sql = "";
                string custName = sheet.OpenCell("CustName").Value.Trim();//获取提交信息，客户名称
                string orderId = sheet.OpenCell("OrderNum").Value;//获取提交信息，订单编号
                string district = sheet.OpenCell("CustDistrict").Value;//获取提交信息，客户所在区域
                string date = DateTime.Now.ToString("yyyy-MM-dd");


                string salesName = sheet.OpenCell("SalesName").Value;//获取提交信息，销售人员姓名
                string amount = sheet.OpenCell("Amount").Value;//获取提交信息，销售金额
                int num = 0;

                if (custName.Trim().Length > 0)
                {
                    if (cid.Trim() != "")
                    {
                        sql = "Update OrderMaster set orderNum = '" + orderId + "',MakerName = '" + "admin"
            + "',CustName='" + custName + "',CustDistrict='" + district + "',SalesName = '" + salesName
            + "' ,Amount= " + amount + " where ID = " + cid;
                }
                    else
                    {
                        sql = "Insert into OrderMaster values(" + (maxId + 1) + ",'" + orderId + "','" + date + "','" + custName + "','"
                            + district + "','" + "admin" + "','" + salesName + "'," + amount + ")";
                    }

                    try
                    {
                        
                        cmd.CommandText = sql;
                        num = cmd.ExecuteNonQuery();//更新客户信息
                    }
                    catch (Exception ex)
                    {

                        strErrHtml += "保存失败，失败原因为：" + ex.Message + "\n";
                    }
                }
                else
                {
                    if (custName.Trim().Length <= 0)
                    {
                        strErrHtml += "请填写订单信息！\n";
                    }
                }

                return num;

            }
        }


    }
