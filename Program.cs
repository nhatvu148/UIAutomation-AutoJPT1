﻿using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;

namespace AutoJPT1
{
    class Program
    {
        static void Main(string[] args)
        {
            WindowsDriver<WindowsElement> sessionJpt;
            AppiumOptions appOptions = new AppiumOptions();
            appOptions.AddAdditionalCapability("app", @"C:\Program Files\TechnoStar\Jupiter-Pre_5.0\DCAD_main.exe");

            sessionJpt = new WindowsDriver<WindowsElement>(
                new Uri("http://127.0.0.1:4723"),
                appOptions
                );

            // IMPORT:
            /* var importCad = sessionJpt.FindElementByName("Import CAD");           
            action.MoveToElement(importCad);
            action.MoveToElement(importCad, importCad.Size.Width/2, importCad.Size.Height/3 + 40).Click().Perform();
            
            sessionJpt.FindElementByName("VRML").Click();

            action.SendKeys(@"V:\Technologies\CS\UIAutomation\ScenarioTestDocs\SampleData\Block\Block.wrl");
            action.SendKeys(Keys.Enter);
            action.Perform();

            sessionJpt.FindElementByName("OK").Click();
            */
            sessionJpt.FindElementByName("Open...").Click();
            Actions action = new Actions(sessionJpt);
            action.SendKeys(@"V:\Technologies\CS\UIAutomation\ScenarioTestDocs\SampleData\Block\Block.jtdb");
            action.SendKeys(Keys.Enter);
            action.Perform();
            
            sessionJpt.FindElementByName("Mesh Cleanup").Click();
            sessionJpt.FindElementByName("Free Edges").Click();

            var allParts = sessionJpt.FindElementByName("All Parts");
            action = new Actions(sessionJpt);
            action.MoveToElement(allParts);
            action.Click();
            // Right Click
            action.ContextClick();
            action.Perform();

            sessionJpt.FindElementByName("Select").Click();
            sessionJpt.FindElementByName("Apply").Click();

            action = new Actions(sessionJpt);
            action.SendKeys(Keys.Enter);
            action.Perform();

            sessionJpt.FindElementByName("Cancel").Click();
            sessionJpt.FindElementByName("Close Gaps").Click();

            sessionJpt.FindElementByAccessibilityId("1001").SendKeys("0.01");
            action = new Actions(sessionJpt);
            action.MoveToElement(allParts);
            action.Click();
            action.ContextClick();
            action.Perform();
            sessionJpt.FindElementByName("Select").Click();
            sessionJpt.FindElementByName("OK").Click();
            

            sessionJpt.FindElementByName("Geometry").Click();
            sessionJpt.FindElementByName("Merge Entities").Click();
            action = new Actions(sessionJpt);
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Perform();
            
            sessionJpt.FindElementByName("Home").Click();
            WindowsElement find = sessionJpt.FindElementByName("Find");
            action = new Actions(sessionJpt);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2, find.Size.Height / 3 + 40).Click();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Perform();

            int [] faceIdArr = { 321, 320, 333, 332, 343, 342, 354, 353, 350, 340, 351, 341, 329, 317, 328 };

            WindowsElement idBox = sessionJpt.FindElementByAccessibilityId("1582");
         
            foreach (int faceId in faceIdArr){
                InputId(faceId, idBox, sessionJpt, find);
            }

            sessionJpt.FindElementByName("OK").Click();



        }
        public static void InputId(int faceId, WindowsElement idBox,
                            WindowsDriver<WindowsElement> sessionJpt, WindowsElement find)
        {
            idBox.SendKeys(Keys.Control + "a" + Keys.Control);
            idBox.SendKeys(Convert.ToString(faceId));
            Actions action5 = new Actions(sessionJpt);
            action5.MoveToElement(find);
            action5.MoveToElement(find, find.Size.Width / 2 - 30, find.Size.Height / 3 - 8).Click().Perform();
        }
    }
}
