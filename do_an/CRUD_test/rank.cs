﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace do_an.CRUD_test
{
    [TestClass]
    public class rank
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
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetAddRankCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void create(string email, string password, string code, string name, string price, string rank, string amount, string date_start, string date_end, string rowValue)
        {
            string actual_result = "";
            bool status = true;
            int row = int.Parse(rowValue);
            Login(email, password);
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                Thread.Sleep(1000);
                if (status)
                {
                    var clickDiscount = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[6]/a[1]/span[1]"));
                    status = clickDiscount != null;
                    if (status)
                    {
                        clickDiscount.Click();
                    }
                    Thread.Sleep(1000);
                    var clickAdd = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[6]/div[1]/div[1]/a[2]"));
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
                    var enterCode = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/input[1]"));
                    status = enterCode != null;
                    if (status)
                    {
                        enterCode.SendKeys(code);
                    }
                    Thread.Sleep(1000);
                    var enterPrice = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/input[1]"));
                    status = enterPrice != null;
                    if (status)
                    {
                        enterPrice.SendKeys(price);
                    }
                    Thread.Sleep(1000);
                    var selectRank = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[4]/select[1]"));
                    Thread.Sleep(1000);
                    status = selectRank != null;
                    if (status)
                    {
                        var selectElementRank = new SelectElement(selectRank);
                        selectElementRank.SelectByValue(rank);
                        selectRank.Click();
                    }
                    Thread.Sleep(1000);
                    var enterAmount = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/input[1]"));
                    status = enterAmount != null;
                    if (status)
                    {
                        enterAmount.SendKeys(amount);
                    }
                    Thread.Sleep(1000);
                    var enterdateStart = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/input[1]"));
                    status = enterdateStart != null;
                    if (status)
                    {
                        enterdateStart.SendKeys(date_start);
                    }
                    Thread.Sleep(1000);
                    var enterdateEnd = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/input[1]"));
                    status = enterdateEnd != null;
                    if (status)
                    {
                        enterdateEnd.SendKeys(date_end);
                    }
                    var add = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                    status = add != null;
                    if (status)
                    {
                        add.Click();
                    }
                    var searchAgain = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = searchAgain != null;
                    if (status)
                    {
                        searchAgain.SendKeys(name);
                        Thread.Sleep(1000);
                        try
                        {
                            var searchCategory = driver.FindElement(By.ClassName("sorting_1"));
                            status = searchCategory != null;
                            if (status)
                            {
                                actual_result = searchCategory.Text;
                            }
                        }
                        catch (Exception ex)
                        {
                            var dataempty = driver.FindElement(By.ClassName("dataTables_empty"));
                            Thread.Sleep(1000);
                            status = dataempty != null;
                            if (status)
                            {
                                actual_result = dataempty.Text;
                            }
                            Thread.Sleep(1000);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                try
                {
                    string validateName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/span[1]")).Text;
                    if (validateName != null)
                    {
                        actual_result = validateName;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    Thread.Sleep(1000);
                    string validateCode = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/span[1]")).Text;
                    if (validateCode != null)
                    {
                        actual_result = actual_result + validateCode;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    Thread.Sleep(1000);
                    string validatePrice = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[3]/span[1]")).Text;
                    if (validatePrice != null)
                    {
                        actual_result = actual_result + validatePrice;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    Thread.Sleep(1000);
                    string validateAmount = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[5]/span[1]")).Text;
                    if (validateAmount != null)
                    {
                        actual_result = actual_result + validateAmount;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
            }
            AddRankExcelResult(actual_result, row);
        }
        [TestMethod]
        [DataTestMethod]
        //[DataRow(CRUD_data.Category.update.Consts.name, CRUD_data.Category.update.Consts.description, CRUD_data.Category.update.Consts.newName)]
        [DynamicData(nameof(GetUpdateRankCredentialsFromExcel), DynamicDataSourceType.Method)]

        public void Update(string email, string password, string name, string newName, string description, string rowValue)
        {
            string actual_result = "";
            bool status = true;
            int row = int.Parse(rowValue);
            Login(email, password);
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
                    var clickCate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[3]/a[1]/span[1]"));
                    status = clickCate != null;
                    if (status)
                    {
                        clickCate.Click();
                    }
                    Thread.Sleep(1000);
                    var clickShow = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[3]/div[1]/div[1]/a[1]"));
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
                        try
                        {
                            var searchAgain = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                            status = searchAgain != null;
                            if (status)
                            {
                                searchAgain.SendKeys(name);
                                Thread.Sleep(1000);
                                try
                                {
                                    var searchCategory = driver.FindElement(By.ClassName("sorting_1"));
                                    status = searchCategory != null;
                                    if (status)
                                    {
                                        actual_result = searchCategory.Text;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    var dataempty = driver.FindElement(By.ClassName("dataTables_empty"));
                                    Thread.Sleep(1000);
                                    status = dataempty != null;
                                    if (status)
                                    {
                                        actual_result = dataempty.Text;
                                    }
                                    Thread.Sleep(1000);
                                }
                            }
                        }
                        catch
                        {
                            var dataempty = driver.FindElement(By.ClassName("dataTables_empty"));
                            Thread.Sleep(1000);
                            status = dataempty != null;
                            if (status)
                            {
                                actual_result = dataempty.Text;
                            }
                            Thread.Sleep(1000);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    string validateName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/span[1]")).Text;
                    if (validateName != null)
                    {
                        actual_result = validateName;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    Thread.Sleep(1000);
                    string validateDescription = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[2]/span[1]")).Text;
                    if (validateDescription != null)
                    {
                        actual_result = actual_result + validateDescription;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
            }
            UpdateRankExcelResult(actual_result, row);
        }
        [TestMethod]
        [DataTestMethod]
        //[DataRow(CRUD_data.Category.delete.Consts.name)]
        [DynamicData(nameof(GetDeleteRankCredentialsFromExcel), DynamicDataSourceType.Method)]

        public void delete(string email, string password, string name, string rowValue)
        {
            string actual_result = "";
            int row = int.Parse(rowValue);
            Login(email, password);
            bool status = true;
            try
            {

                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
                    var clickCate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[3]/a[1]/span[1]"));
                    status = clickCate != null;
                    if (status)
                    {
                        clickCate.Click();
                    }
                    Thread.Sleep(1000);
                    var clickShow = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[3]/div[1]/div[1]/a[1]"));
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
                        //var dataempty = driver.FindElement(By.ClassName("dataTables_empty"));
                        var dataempty = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]"));
                        Thread.Sleep(1000);
                        status = dataempty != null;
                        if (status)
                        {
                            actual_result = dataempty.Text;
                            status = true;
                        }
                        Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception ex)
            {
                actual_result = name;
                driver.Quit();
            }
            DeleteRankExcelResult(actual_result, row);
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetSortRankCredentialsFromExcel), DynamicDataSourceType.Method)]

        public void sort(string email, string password, string sort, string rowValue)
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
                    var clickBrand = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[3]/a[1]/span[1]"));
                    status = clickBrand != null;
                    if (status)
                    {
                        clickBrand.Click();
                    }
                    Thread.Sleep(1000);
                    var clickShow = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[3]/div[1]/div[1]/a[1]"));
                    status = clickShow != null;
                    if (status)
                    {
                        clickShow.Click();
                    }
                    Thread.Sleep(1000);
                    if (sort == "Name")
                    {
                        var clickSortName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[1]"));
                        status = clickSortName != null;
                        if (status)
                        {
                            clickSortName.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Description")
                    {
                        var clickSortDescription = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[2]"));
                        status = clickSortDescription != null;
                        if (status)
                        {
                            clickSortDescription.Click();
                            clickSortDescription.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[2]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Created at")
                    {
                        var clickSortCrea = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[3]"));
                        status = clickSortCrea != null;
                        if (status)
                        {
                            clickSortCrea.Click();
                            clickSortCrea.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[3]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Updated at")
                    {
                        var clickSortUp = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[4]"));
                        status = clickSortUp != null;
                        if (status)
                        {
                            clickSortUp.Click();
                            clickSortUp.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[4]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                }
            }
            catch
            {
                driver.Quit();
            }
            SortRankExcelResult(actual_result, row);
        }
        private static IEnumerable<object[]> GetAddRankCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[5];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 3].Value.ToString();
                    string password = worksheet.Cells[row, 4].Value.ToString();
                    object cellValueName = worksheet.Cells[row, 5].Value;
                    string name = cellValueName != null ? cellValueName.ToString() : string.Empty;
                    object cellValueDescription = worksheet.Cells[row, 6].Value;
                    string description = cellValueDescription != null ? cellValueDescription.ToString() : string.Empty;
                    yield return new string[] { email, password, name, description, rowValue };
                }
            }
        }

        private void AddRankExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[5];

                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 7].Value.ToString();
                if (actual_result == expected)
                {
                    worksheet.Cells[row, 8].Value = actual_result;
                    worksheet.Cells[row, 9].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 8].Value = actual_result;
                    worksheet.Cells[row, 9].Value = "Faild";
                }
                package.Save();
            }
        }
        private static IEnumerable<object[]> GetUpdateRankCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[6];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 3].Value.ToString();
                    string password = worksheet.Cells[row, 4].Value.ToString();
                    object cellValueName = worksheet.Cells[row, 5].Value;
                    string name = cellValueName != null ? cellValueName.ToString() : string.Empty;
                    object cellValueNewName = worksheet.Cells[row, 6].Value;
                    string newname = cellValueNewName != null ? cellValueNewName.ToString() : string.Empty;
                    object cellValueDescription = worksheet.Cells[row, 7].Value;
                    string description = cellValueDescription != null ? cellValueDescription.ToString() : string.Empty;
                    yield return new string[] { email, password, name, newname, description, rowValue };
                }
            }
        }
        private void UpdateRankExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[6];
                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 8].Value.ToString();
                if (actual_result == expected)
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
        private static IEnumerable<object[]> GetDeleteRankCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[7];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    object cellValueEmail = worksheet.Cells[row, 3].Value;
                    string email = cellValueEmail != null ? cellValueEmail.ToString() : string.Empty;
                    object cellValuePassword = worksheet.Cells[row, 4].Value;
                    string password = cellValuePassword != null ? cellValuePassword.ToString() : string.Empty;
                    object cellValueName = worksheet.Cells[row, 5].Value;
                    string name = cellValueName != null ? cellValueName.ToString() : string.Empty;
                    yield return new string[] { email, password, name, rowValue };
                }
            }
        }
        private void DeleteRankExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[7];
                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 6].Value.ToString();
                if (actual_result == expected)
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
        private static IEnumerable<object[]> GetSortRankCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[8];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 3].Value.ToString();
                    string password = worksheet.Cells[row, 4].Value.ToString();
                    object cellValueSort = worksheet.Cells[row, 5].Value;
                    string sort = cellValueSort != null ? cellValueSort.ToString() : string.Empty;
                    yield return new string[] { email, password, sort, rowValue };
                }
            }
        }
        private void SortRankExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[8];
                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 6].Value.ToString();
                if (actual_result == expected)
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
    }
}
