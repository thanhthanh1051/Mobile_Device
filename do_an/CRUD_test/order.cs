using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Support.UI;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace do_an.CRUD_test
{
    [TestClass]
    public class order
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
        }
        public void Logout()
        {
            driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[1]/div/div[3]/div/a[1]/div/span[1]")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[2]/form/div/a")).Click();
            Thread.Sleep(3000);
        }

        public void Status(string email, string password, string statusOrder, string rowValue)
        {
            int row = int.Parse(rowValue);
            string actual_result = "";
            Login(email, password);
            Thread.Sleep(1000);
            bool status = true;
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
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
                        actual_result = "Pass";
                    }
                    Thread.Sleep(1000);
                }
               }
            catch(Exception ex)
            {
                actual_result = "Faild";
                driver.Quit();
            }
            StatusOrderExcelResult(actual_result, row);
        }
        private static IEnumerable<object[]> GetStatusOrderCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[12];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 3].Value.ToString();
                    string password = worksheet.Cells[row, 4].Value.ToString();
                    object statusOrder = worksheet.Cells[row, 5].Value;
                    string status = statusOrder != null ? statusOrder.ToString() : string.Empty;
                    yield return new string[] { email, password, status, rowValue };
                }
            }
        }
        private void StatusOrderExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[12];
                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 6].Value.ToString();
                if (actual_result == expected)
                {
                    worksheet.Cells[row, 7].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 7].Value = "Faild";
                }
                package.Save();
            }
        }
    }
}
