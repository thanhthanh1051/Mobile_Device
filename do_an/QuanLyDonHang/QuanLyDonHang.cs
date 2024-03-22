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
using SeleniumExtras.WaitHelpers;
using System.Windows.Input;

namespace do_an.QuanLyDonHang
{
    [TestClass]
    public class QuanLyDonHang
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
        public void addOrder(string namePro_type, string amountPro)
        {
            bool status = true;
            var clickSearch = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/div[1]/input[1]"));
            status = clickSearch != null;
            try
            {
                    if (status)
                    {
                        clickSearch.SendKeys(namePro_type);
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
                    var amountCart = driver.FindElement(By.ClassName("icon-amount-orders"));
                    status = amountCart != null;
                    if (status)
                    {
                        string amount = amountCart.Text;
                        if (amount == amountPro)
                        {
                            amountCart.Click();
                        }
                        Thread.Sleep(1000);
                        var namePro = driver.FindElement(By.CssSelector("div[class='product-name'] h5"));
                        status = namePro != null;
                        if (status)
                        {
                            string name = namePro.Text;
                            if (name == namePro_type)
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
                                    clickChangeInfo.SendKeys("");
                                }
                                Thread.Sleep(1000);
                                var enterAddress = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/div[6]/input[1]"));
                                status = enterAddress != null;
                                if (status)
                                {
                                    enterAddress.SendKeys("");
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
                            }
                        }
                    }
            }catch (Exception ex)
            {
                driver.Quit();
            }
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetCreateOrderCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void CreateOrder(string email, string password, string namePro_type, string amountPro)
        {
            string actual_result = "";
            bool status = true;
            Login(email, password);
            var enterSearch = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/div[1]/input[1]"));
            status = enterSearch != null;
            try
            {
                if (status)
                {
                    enterSearch.SendKeys(namePro_type);
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
                    string amount = amountCart.Text;
                    if (amount == amountPro)
                    {
                        amountCart.Click();
                    }
                    Thread.Sleep(1000);
                    var namePro = driver.FindElement(By.CssSelector("div[class='product-name'] h5"));
                    status = namePro != null;
                    if (status)
                    {
                        string name = namePro.Text;
                        if (name == namePro_type)
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
                                clickChangeInfo.SendKeys("hello");
                            }
                            Thread.Sleep(1000);
                            var enterAddress = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/form[1]/div[6]/input[1]"));
                            status = enterAddress != null;
                            if (status)
                            {
                                enterAddress.SendKeys("abc");
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
                        }else
                        {
                            actual_result = "Thanh toán thất bại";
                        }
                    }else
                    {
                        actual_result = "Thanh toán thất bại";
                    }
                }
            }
            catch (Exception ex)
            {
                actual_result = "Thanh toán thất bại";
                UpdateOrderExcelResult(actual_result);
                driver.Quit();
            }
            UpdateOrderExcelResult(actual_result);
        }
        private static IEnumerable<object[]> GetCreateOrderCredentialsFromExcel()
        {
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 2].Value.ToString();
                    string password = worksheet.Cells[row, 3].Value.ToString();
                    string nameProduct = worksheet.Cells[row, 4].Value.ToString();
                    string amountProduct = worksheet.Cells[row, 5].Value.ToString();
                    yield return new string[] { email, password, nameProduct, amountProduct };
                }
            }
        }
        private void UpdateOrderExcelResult(string actual_result)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string expected = worksheet.Cells[row, 6].ToString();
                    worksheet.Cells[row, 7].Value = actual_result;
                    if(expected.Equals(actual_result))
                    {
                        worksheet.Cells[row, 8].Value = "Pass";
                    }
                    else
                    {
                        worksheet.Cells[row, 8].Value = "Faild";
                    }
                    
                }
                package.Save();
            }
        }
        [TestCleanup]
        public void clear()
        {
            driver.Close();
            driver.Quit();
        }
    }
}
