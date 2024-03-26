using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using do_an.CRUD_test;

namespace do_an.DashBoard
{
    [TestClass]
    public class DashBoard
    {
        IWebDriver driver = new ChromeDriver();
        [TestInitialize]
        public void Init()
        {
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);

            driver.Url = "http://localhost:81/";
            driver.Navigate();
            Thread.Sleep(1000);
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
        }
        public void Login(string email, string password)
        {
            bool status = true;
            try
            {
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
                        enterEmail.SendKeys(email);
                    }
                    Thread.Sleep(2000);
                    var enterPassword = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/input[2]"));
                    Thread.Sleep(1000);
                    status = enterPassword != null;
                    if (status)
                    {
                        enterPassword.SendKeys(password);
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
            //Assert.IsTrue(status);
        }
        public void Logout()
        {
            driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[1]/div/div[3]/div/a[1]/div/span[1]")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[2]/form/div/a")).Click();
            Thread.Sleep(3000);
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetManagerOrderCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void Dashboard(string emailCus, string passwordCus, string emailAd, string passAd, string namePro, string amount, string phone, string address, string statusOrder, string rowValue)
        {
            int row = int.Parse(rowValue);
            string actual_result = "";
            bool status = true;
            Login(emailCus, passwordCus);
            var enterSearch = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/div[1]/input[1]"));
            status = enterSearch != null;
            try
            {
                if (status)
                {
                    enterSearch.SendKeys(namePro);
                }
                Thread.Sleep(1000);
                var clickSearch = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/div[1]/button[1]/i[1]"));
                status = clickSearch != null;
                if (status)
                {
                    clickSearch.Click();
                }
                Thread.Sleep(1000);
                var itemCart = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[1]/div[1]/img[1]"));
                status = itemCart != null;
                if (status)
                {
                    itemCart.Click();
                }
                Thread.Sleep(1000);
                var addCart = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/a[1]"));
                status = addCart != null;
                if (status)
                {
                    addCart.Click();
                }
                Thread.Sleep(1000);
                var clickOkCart = driver.FindElement(By.XPath("/html[1]/body[1]/div[2]/div[1]/div[6]/button[1]"));
                status = clickOkCart != null;
                if (status)
                {
                    clickOkCart.Click();
                }
                Thread.Sleep(1000);
                //var amountCart = driver.FindElement(By.ClassName("icon-amount-orders"));
                var amountCart = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[3]/div[1]/a[3]/div[1]/span[2]"));
                status = amountCart != null;
                if (status)
                {
                    string amountpro = amountCart.Text;
                    if (amountpro == amount)
                    {
                        amountCart.Click();
                    }
                    Thread.Sleep(1000);
                    var nameproduct = driver.FindElement(By.CssSelector("div[class='product-name'] h5"));
                    status = nameproduct != null;
                    if (status)
                    {
                        string name = nameproduct.Text;
                        if (name == namePro)
                        {
                            Thread.Sleep(1000);
                            var clickCheckOut = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[2]/div[2]/a[1]"));
                            status = clickCheckOut != null;
                            if (status)
                            {
                                clickCheckOut.Click();
                            }
                            Thread.Sleep(1000);
                            var clickChangeInfo = driver.FindElement(By.XPath("/html[1]/body[1]/div[6]/div[1]/div[6]/button[1]/a[1]"));
                            status = clickChangeInfo != null;
                            if (status)
                            {
                                clickChangeInfo.Click();
                            }
                            Thread.Sleep(1000);
                            var enterPhone = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/div[4]/input[1]"));
                            status = enterPhone != null;
                            if (status)
                            {
                                clickChangeInfo.SendKeys(phone);
                            }
                            Thread.Sleep(1000);
                            var enterAddress = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/div[6]/input[1]"));
                            status = enterAddress != null;
                            if (status)
                            {
                                enterAddress.SendKeys(address);
                            }
                            Thread.Sleep(1000);
                            var clickSubmit = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/button[1]"));
                            status = clickSubmit != null;
                            if (status)
                            {
                                clickSubmit.Click();
                            }
                            Thread.Sleep(1000);
                            var checkOut = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[2]/div[2]/a[1]"));
                            status = checkOut != null;
                            if (status)
                            {
                                clickCheckOut.Click();
                            }
                            Thread.Sleep(1000);
                            var susscess = driver.FindElement(By.ClassName("swal2-title"));
                            status = susscess != null;
                            if (status)
                            {
                                string sc = susscess.Text;
                                var clickOk = driver.FindElement(By.XPath("/html[1]/body[1]/div[6]/div[1]/div[6]/button[1]"));
                                clickOk.Click();
                                actual_result = sc;
                            }
                            Login("admin123@gmail.com", "admin123");
                            var clickOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/a[1]/span[1]"));
                            status = clickOrder != null;
                            if (status)
                            {
                                clickOrder.Click();
                            }
                            Thread.Sleep(1000);
                            var clickShow = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/div[1]/div[1]/a[1]"));
                            status = clickShow != null;
                            if (status)
                            {
                                clickShow.Click();
                            }
                            Thread.Sleep(1000);
                            var clickUpdate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[9]/a[1]"));
                            status = clickUpdate != null;
                            if (status)
                            {
                                clickUpdate.Click();
                            }
                            Thread.Sleep(1000);
                            var clickStatus = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[3]/select[1]"));
                            status = clickStatus != null;
                            if (status)
                            {
                                var selectElement = new SelectElement(clickStatus);
                                selectElement.SelectByValue(statusOrder);
                                clickStatus.Click();
                            }
                            Thread.Sleep(1000);
                            var update = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                            status = update != null;
                            if (status)
                            {
                                update.Click();
                            }
                            Thread.Sleep(1000);
                            var clickDashboard = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[1]/a[1]/span[1]"));
                            status = clickDashboard != null;
                            if (status)
                            {
                                clickDashboard.Click();
                            }
                            var checkAnnual = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]"));
                            status = checkAnnual != null;
                            if (status)
                            {
                                actual_result = checkAnnual.Text;
                            }
                            Thread.Sleep(1000);
                            var checkMonthly = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]"));
                            status = checkMonthly != null;
                            if (status)
                            {
                                if(checkMonthly.Text == checkAnnual.Text)
                                {
                                    actual_result = checkMonthly.Text;
                                }
                            }
                        }
                        else
                        {
                            actual_result = "Thanh toán thất bại";
                        }
                    }
                    else
                    {
                        actual_result = "Thanh toán thất bại";
                    }
                }
            }
            catch (Exception ex)
            {
                actual_result = "not right";
                driver.Quit();
            }
            ManagerOrderExcelResult(actual_result, row);
        }

        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetManagerOrderCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void ThongKeBaoCao(string email, string password, string statusOrder, string rowValue)
        {
            //string total = "";
            string actual_result = "";
            bool status = true;
            int row = int.Parse(rowValue);
            //int amount = int.Parse(amountOrder);
            Login(email, password);
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                Thread.Sleep(1000);
                if (status)
                {
                    //var clickOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/a[1]"));
                    //status = clickOrder != null;
                    //if (status)
                    //{
                    //    clickOrder.Click();
                    //}
                    //Thread.Sleep(1000);
                    //var showOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/div[1]/div[1]/a[1]"));
                    //status = showOrder != null;
                    //if(status)  
                    //{
                    //    showOrder.Click();
                    //}
                    //Thread.Sleep(1000);
                    //for(int i = 2; i <= amount+1; i++)
                    //{
                    //    var takeTotal = driver.FindElement(By.XPath("td:nth-child(" + i + ")"));
                    //    status = takeTotal != null;
                    //    if (status)
                    //    {
                    //        total = takeTotal.Text;
                    //    }
                    //}
                    //Thread.Sleep(1000);
                    var clickOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/a[1]"));
                    status = clickOrder != null;
                    if (status)
                    {
                        clickOrder.Click();
                    }
                    Thread.Sleep(1000);
                    var showOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/div[1]/div[1]/a[1]"));
                    status = showOrder != null;
                    if (status)
                    {
                        showOrder.Click();
                    }
                    Thread.Sleep(1000);
                    var clickUpdate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[9]/a[1]"));
                    status = clickUpdate != null;
                    if (status)
                    {
                        clickUpdate.Click();
                    }
                    Thread.Sleep(1000);
                    var selectStatus = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[3]/select[1]"));
                    Thread.Sleep(1000);
                    status = selectStatus != null;
                    if (status)
                    {
                        var selectElementOrder = new SelectElement(selectStatus);
                        selectElementOrder.SelectByValue(statusOrder);
                        selectStatus.Click();
                    }
                    Thread.Sleep(1000);
                    var clickChange = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                    status = clickChange != null;
                    if (status)
                    {
                        clickChange.Click();
                    }
                    Thread.Sleep(1000);
                    var clickDashboard = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[1]/a[1]/span[1]"));
                    status = clickDashboard != null;
                    if (status)
                    {
                        clickDashboard.Click();
                    }
                    Thread.Sleep(1000);
                    var annual = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]"));
                    status = annual != null;
                    if (status)
                    {
                        actual_result = annual.Text;
                    }
                }
            }
            catch(Exception ex)
            {
                driver.Quit();
            }
            ManagerOrderExcelResult(actual_result, row);
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetManagerOrder2CredentialsFromExcel), DynamicDataSourceType.Method)]
        public void ThongKeBaoCao2(string email, string password, string phone1, string phone2, string statusOrder, string rowValue)
        {
            //string total = "";
            string actual_result = "";
            bool status = true;
            int row = int.Parse(rowValue);
            //int amount = int.Parse(amountOrder);
            Login(email, password);
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                Thread.Sleep(1000);
                if (status)
                {
                    //var clickOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/a[1]"));
                    //status = clickOrder != null;
                    //if (status)
                    //{
                    //    clickOrder.Click();
                    //}
                    //Thread.Sleep(1000);
                    //var showOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/div[1]/div[1]/a[1]"));
                    //status = showOrder != null;
                    //if(status)  
                    //{
                    //    showOrder.Click();
                    //}
                    //Thread.Sleep(1000);
                    //for(int i = 2; i <= amount+1; i++)
                    //{
                    //    var takeTotal = driver.FindElement(By.XPath("td:nth-child(" + i + ")"));
                    //    status = takeTotal != null;
                    //    if (status)
                    //    {
                    //        total = takeTotal.Text;
                    //    }
                    //}
                    //Thread.Sleep(1000);
                    var clickOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/a[1]"));
                    status = clickOrder != null;
                    if (status)
                    {
                        clickOrder.Click();
                    }
                    Thread.Sleep(1000);
                    var showOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/div[1]/div[1]/a[1]"));
                    status = showOrder != null;
                    if (status)
                    {
                        showOrder.Click();
                    }
                    Thread.Sleep(1000);
                    var searchNumberORder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = searchNumberORder != null;
                    if(status)
                    {
                        searchNumberORder.SendKeys(phone1);
                    }
                    Thread.Sleep(1000);
                    var clickUpdate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[9]/a[1]"));
                    status = clickUpdate != null;
                    if (status)
                    {
                        clickUpdate.Click();
                    }
                    Thread.Sleep(1000);
                    var selectStatus = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[3]/select[1]"));
                    Thread.Sleep(1000);
                    status = selectStatus != null;
                    if (status)
                    {
                        var selectElementOrder = new SelectElement(selectStatus);
                        selectElementOrder.SelectByValue(statusOrder);
                        selectStatus.Click();
                    }
                    Thread.Sleep(1000);
                    var clickChange = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                    status = clickChange != null;
                    if (status)
                    {
                        clickChange.Click();
                    }
                    Thread.Sleep(1000);
                    Thread.Sleep(1000);
                    var searchNumberORder2 = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = searchNumberORder2 != null;
                    if (status)
                    {
                        searchNumberORder2.SendKeys(phone2);
                    }
                    Thread.Sleep(1000);
                    var clickUpdate2 = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[9]/a[1]"));
                    status = clickUpdate2 != null;
                    if (status)
                    {
                        clickUpdate2.Click();
                    }
                    Thread.Sleep(1000);
                    var selectStatus2 = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[3]/select[1]"));
                    Thread.Sleep(1000);
                    status = selectStatus2 != null;
                    if (status)
                    {
                        var selectElementOrder = new SelectElement(selectStatus2);
                        selectElementOrder.SelectByValue(statusOrder);
                        selectStatus2.Click();
                    }
                    Thread.Sleep(1000);
                    var clickChange2 = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                    status = clickChange2 != null;
                    if (status)
                    {
                        clickChange2.Click();
                    }
                    Thread.Sleep(1000);

                    var clickDashboard = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[1]/a[1]/span[1]"));
                    status = clickDashboard != null;
                    if (status)
                    {
                        clickDashboard.Click();
                    }
                    Thread.Sleep(1000);
                    var annual = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]"));
                    status = annual != null;
                    if (status)
                    {
                        actual_result = annual.Text;
                    }
                }
            }
            catch (Exception ex)
            {
                driver.Quit();
            }
            ManagerOrder2ExcelResult(actual_result, row);
        }

        private static IEnumerable<object[]> GetManagerOrderCredentialsFromExcel()
        {
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["ThongKeBaoCao"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string emailCus = worksheet.Cells[row, 3].Value.ToString();
                    string passwordCus = worksheet.Cells[row, 4].Value.ToString();
                    string statusOrder = worksheet.Cells[row, 5].Value.ToString();
                    yield return new string[] { emailCus, passwordCus, statusOrder, rowValue };
                }
            }
        }
        private void ManagerOrderExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["ThongKeBaoCao"];

                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 6].Value.ToString();
                if (expected == actual_result)
                {
                    worksheet.Cells[row, 7].Value = actual_result;
                    worksheet.Cells[row, 8].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 7].Value = actual_result;
                    worksheet.Cells[row, 8].Value = "Faild";
                }
                package.Save();
            }
        }
        private static IEnumerable<object[]> GetManagerOrder2CredentialsFromExcel()
        {
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["ThongKeBaoCao2"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string emailCus = worksheet.Cells[row, 3].Value.ToString();
                    string passwordCus = worksheet.Cells[row, 4].Value.ToString();
                    string phone1 = worksheet.Cells[row, 5].Value.ToString();
                    string phone2 = worksheet.Cells[row, 6].Value.ToString();
                    string statusOrder = worksheet.Cells[row, 7].Value.ToString();
                    yield return new string[] { emailCus, passwordCus, phone1, phone2, statusOrder, rowValue };
                }
            }
        }
        private void ManagerOrder2ExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["ThongKeBaoCao2"];

                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 8].Value.ToString();
                if (expected == actual_result)
                {
                    worksheet.Cells[row, 9].Value = actual_result;
                    worksheet.Cells[row, 10].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 9].Value = actual_result;
                    worksheet.Cells[row, 10].Value = "Faild";
                }
                package.Save();
            }
        }
        [TestCleanup]
        public void clear()
        {
            driver.Quit();
        }
    }
}

