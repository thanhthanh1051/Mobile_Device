using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Net.Configuration;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Support.UI;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using System.IO;
using Newtonsoft.Json;


namespace do_an.CRUD_test
{
    [TestClass]
    public class brand
    {
        IWebDriver driver = new ChromeDriver();
        string fNumber = null;
        string sNumber = null;
        string f1Number = null;
        string f2Number = null;
        [TestInitialize]
        public void Test_Login()
        {
            bool status = true;
            try
            {
                driver.Manage().Window.Maximize();
                Thread.Sleep(1000);

                driver.Url = "http://localhost:81/";
                driver.Navigate();
                status = driver != null;
                if (status)
                {
                    Thread.Sleep(2000);
                    var iconLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[3]/div[1]/a[1]/div[1]/span[1]"));
                    Thread.Sleep(1000);
                    status = iconLogin != null;
                    if (status)
                    {
                        iconLogin.Click();
                    }
                    Thread.Sleep(1000);
                    var enterEmail = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/input[1]"));
                    Thread.Sleep(1000);
                    status = enterEmail != null;
                    if (status)
                    {
                        enterEmail.SendKeys("admin123@gmail.com");
                    }
                    Thread.Sleep(2000);
                    var enterPassword = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/input[2]"));
                    Thread.Sleep(1000);
                    status = enterPassword != null;
                    if (status)
                    {
                        enterPassword.SendKeys("admin123");
                    }
                    Thread.Sleep(2000);
                    var clickLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/button[1]"));
                    Thread.Sleep(1000);
                    status = clickLogin != null;
                    if (status)
                    {
                        clickLogin.Click();
                    }
                    Thread.Sleep(2000);
                }
            }
            catch (Exception ex)
            {
                Assert.IsFalse(status);
                driver.Quit();
            }
            Assert.IsTrue(status);
        }
        [TestMethod]
        [DataTestMethod]
        [DataRow(CRUD_data.Brand.create.Consts.name, CRUD_data.Brand.create.Consts.description)]
        public void create(string name, string description)
        {
            bool status = true;
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                Thread.Sleep(1000);
                if (status)
                {
                    var clickCreate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[4]/a[1]/span[1]"));
                    status = clickCreate != null;
                    if (status)
                    {
                        clickCreate.Click();
                    }
                    Thread.Sleep(1000);
                    var clickAdd = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[4]/div[1]/div[1]/a[2]"));
                    status = clickAdd != null;
                    if (status)
                    {
                        clickAdd.Click();
                    }
                    Thread.Sleep(1000);
                    var enterName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/input[1]"));
                    status = enterName != null;
                    if (status)
                    {
                        enterName.SendKeys(name);
                    }
                    Thread.Sleep(1000);
                    var enterDescription = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/input[1]"));
                    status = enterDescription != null;
                    if (status)
                    {
                        enterDescription.SendKeys(description);
                    }
                    Thread.Sleep(1000);
                    var add = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                    status = add != null;
                    if (status)
                    {
                        add.Click();
                    }
                    Thread.Sleep(1000);
                    var checkSreach = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = checkSreach != null;
                    if (status)
                    {
                        checkSreach.SendKeys(name);
                        Thread.Sleep(1000);
                        var searchCode = driver.FindElement(By.ClassName("sorting_1"));
                        string checkName = searchCode.Text;
                        if (checkName == name)
                        {
                            status = true;
                        }
                        else
                        {
                            status = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Assert.IsFalse(status);
                driver.Close();
                driver.Quit();
            }
            Assert.IsTrue(status);
            driver.Close();
        }
        [TestMethod]
        [DataTestMethod]
        [DataRow(CRUD_data.Brand.update.Consts.name, CRUD_data.Brand.update.Consts.description, CRUD_data.Brand.update.Consts.newName)]
        public void Update(string name, string description, string newName)
        {
            bool status = true;
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
                    var clickBrand = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[4]/a[1]/span[1]"));
                    status = clickBrand != null;
                    if (status)
                    {
                        clickBrand.Click();
                    }
                    Thread.Sleep(1000);
                    var clickShow = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[4]/div[1]/div[1]/a[1]"));
                    status = clickShow != null;
                    if (status)
                    {
                        clickShow.Click();
                    }
                    Thread.Sleep(1000);
                    var clickSearch = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = clickSearch != null;
                    if (status)
                    {
                        clickSearch.SendKeys(name);
                    }
                    Thread.Sleep(1000);
                    var searchName = driver.FindElement(By.ClassName("sorting_1"));
                    string check = searchName.Text;
                    if (check == name)
                    {
                        var update = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[5]/a[1]"));
                        status = update != null;
                        if (status)
                        {
                            update.Click();
                        }
                        Thread.Sleep(1000);
                        var enterName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/input[1]"));
                        status = enterName != null;
                        if (status)
                        {
                            enterName.Clear();
                            enterName.SendKeys(newName);
                        }
                        Thread.Sleep(1000);
                        var enterDescription = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/input[1]"));
                        status = (enterDescription != null);
                        if (status)
                        {
                            enterDescription.Clear();
                            enterDescription.SendKeys(description);
                        }
                        Thread.Sleep(1000);
                        var updateChange = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                        status = (updateChange != null);
                        if (status)
                        {
                            updateChange.Click();
                        }
                        Thread.Sleep(1000);
                        var searchAgain = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                        status = searchAgain != null;
                        if (status)
                        {
                            searchAgain.SendKeys(newName);
                        }
                        Thread.Sleep(1000);
                        var dataempty = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]"));
                        Thread.Sleep(1000);
                        status = dataempty != null;
                        if (status)
                        {
                            status = true;
                        }
                        Thread.Sleep(1000);
                    }
                }

            }
            catch (Exception ex)
            {
                Assert.IsFalse(status);
                driver.Close();
                driver.Quit();
            }
            Assert.IsTrue(status);
            driver.Close();
        }

        [TestMethod]
        [DataTestMethod]
        [DataRow(CRUD_data.Brand.delete.Consts.name)]
        public void delete(string name)
        {
            bool status = true;
            try
            {

                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
                    var clickCate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[4]/a[1]/span[1]"));
                    status = clickCate != null;
                    if (status)
                    {
                        clickCate.Click();
                    }
                    Thread.Sleep(1000);
                    var clickShow = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[4]/div[1]/div[1]/a[1]"));
                    status = clickShow != null;
                    if (status)
                    {
                        clickShow.Click();
                    }
                    Thread.Sleep(1000);
                    var clickSearch = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = clickSearch != null;
                    if (status)
                    {
                        clickSearch.SendKeys(name);
                    }
                    Thread.Sleep(1000);
                    var searchName = driver.FindElement(By.ClassName("sorting_1"));
                    string check = searchName.Text;
                    if (check == name)
                    {
                        var dele = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[6]/a[1]"));
                        status = dele != null;
                        if (status)
                        {
                            dele.Click();
                        }
                        Thread.Sleep(1000);
                        var searchAgain = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                        status = searchAgain != null;
                        if (status)
                        {
                            searchAgain.SendKeys(name);
                        }
                        Thread.Sleep(1000);
                        var dataempty = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]"));
                        Thread.Sleep(1000);
                        status = dataempty != null;
                        if (status)
                        {
                            status = true;
                        }
                        Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception ex)
            {
                Assert.IsFalse(status);
                driver.Close();
                driver.Quit();
            }
            Assert.IsTrue(status);
            driver.Close();
        }

        //void ReadFileExel(string filePath)
        //{
        //    // Đọc dữ liệu từ các ô trong worksheet
        //    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        //    Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

        //    Microsoft.Office.Interop.Excel._Worksheet worksheet = workbook.Sheets[1];
        //    Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

        //    for (int row = 2; row <= range.Rows.Count; row++)
        //    {
        //        for (int col = 1; col <= range.Columns.Count; col++)
        //        {
        //            if(row == 3)
        //            {
        //                if(col >= 2)
        //                {
        //                    string cellValue = range.Cells[row, col].Value2?.ToString();
        //                    if (cellValue != null)
        //                    {
        //                        if (col == 2)
        //                        {
        //                            fNumber = cellValue;
        //                        }
        //                        else if(col == 3)
        //                        {
        //                            f1Number = cellValue;
        //                        }
        //                        else if(col == 4)
        //                        {
        //                            f2Number = cellValue;
        //                        }
        //                        else
        //                        {
        //                            sNumber = cellValue;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    workbook.Close();
        //    excelApp.Quit();
        //}

        //static void WriteToColumnC(string filePath, Boolean result)
        //{
        //    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
        //    Microsoft.Office.Interop.Excel._Worksheet worksheet = workbook.Sheets[1];
        //    Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;
        //    for (int row = 2; row <= range.Rows.Count; row++)
        //    {
        //        // Lấy giá trị từ ô đầu tiên

        //        if (result)
        //        {
        //            // Ghi giá trị chuỗi cộng vào cột thứ 3 (cột C)
        //            range.Cells[3, 9].Value2 = "Passed";
        //        }
        //        else
        //        {
        //            range.Cells[3, 9].Value2 = "Failed";
        //        }
        //    }
        //    workbook.Close();
        //    excelApp.Quit();
        //}
        [TestCleanup]
        public void clear()
        {
            driver.Quit();
        }

        //[DynamicData(nameof(GetLoginCredentialsFromExcel), DynamicDataSourceType.Method)]
        //private static IEnumerable<object[]> GetLoginCredentialsFromExcel()
        //{
        //    string filePath = @"D:\Huflit\Năm 3\Học kì 2\Bảo đảm chất lượng phần mềm\Data\dataTH.xlsx";
        //    using (var package = new ExcelPackage(new FileInfo(filePath)))
        //    {
        //        var worksheet = package.Workbook.Worksheets[0];
        //        int rowCount = worksheet.Dimension.Rows;

        //        for (int row = 4; row <= rowCount; row++)
        //        {
        //            string name = worksheet.Cells[row, 1].Value.ToString();
        //            string email = worksheet.Cells[row, 2].Value.ToString();
        //            string subject = worksheet.Cells[row, 3].Value.ToString();
        //            string message = worksheet.Cells[row, 4].Value.ToString();
        //            yield return new string[] { name, email, subject, message };
        //        }
        //    }
        //}

        //private void UpdateExcelResult(string name, string email, string subject, string message, string result)
        //{
        //    string filePath = @"D:\Huflit\Năm 3\Học kì 2\Bảo đảm chất lượng phần mềm\Data\dataTH.xlsx";
        //    FileInfo file = new FileInfo(filePath);

        //    using (ExcelPackage package = new ExcelPackage(file))
        //    {
        //        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

        //        int rowCount = worksheet.Dimension.Rows;

        //        for (int row = 4; row <= rowCount; row++)
        //        {
        //            string Name = worksheet.Cells[row, 1].Value.ToString();
        //            string Email = worksheet.Cells[row, 2].Value.ToString();
        //            string Subject = worksheet.Cells[row, 3].Value.ToString();
        //            string Message = worksheet.Cells[row, 4].Value.ToString();

        //            if (Name == name && Email == email && Subject == subject && Message == message)
        //            {
        //                worksheet.Cells[row, 5].Value = result;
        //                break;
        //            }
        //        }
        //        package.Save();
        //    }
        //}
    }
}
