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
            System.Threading.Thread.Sleep(3000);
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
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Perform();

            sessionJpt.FindElementByName("Apply").Click();

            action = new Actions(sessionJpt);
            action.SendKeys(Keys.Enter);
            action.Perform();

            sessionJpt.FindElementByName("Cancel").Click();
            sessionJpt.FindElementByName("Close Gaps").Click();

            action = new Actions(sessionJpt);
            action.MoveToElement(allParts);
            action.Click();
            action.ContextClick();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Perform();

            typeInBox(1001, 0.01, sessionJpt);

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

            int[] faceIdArr = { 321, 317, 316, 320, 333, 343, 354, 353, 342, 332, 328, 340, 350, 351, 341, 329 };

            WindowsElement idBox = sessionJpt.FindElementByAccessibilityId("1582");

            foreach (int faceId in faceIdArr)
            {
                InputId(faceId, idBox, action, sessionJpt, find);
            }

            sessionJpt.FindElementByName("OK").Click();

            sessionJpt.FindElementByName("Meshing").Click();

            WindowsElement localSettings = sessionJpt.FindElementByName("Local Settings");
            action = new Actions(sessionJpt);
            action.MoveToElement(localSettings);
            action.MoveToElement(localSettings, localSettings.Size.Width / 2 - 50, localSettings.Size.Height / 2)
                .Click();
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Perform();

            action = new Actions(sessionJpt);
            action.SendKeys(Keys.Down);
            action.SendKeys(Keys.Enter);
            action.Perform();

            sessionJpt.FindElementByName("Home").Click();

            InputId(708, idBox, action, sessionJpt, find);
            InputId(698, idBox, action, sessionJpt, find);

            sessionJpt.FindElementByName("Mesh Size").Click();
            typeInBox(1006, 4, sessionJpt);
            typeInBox(1007, 1, sessionJpt);
            typeInBox(1008, 8, sessionJpt);

            sessionJpt.FindElementByName("Apply").Click();

            InputId(1, idBox, action, sessionJpt, find);
            InputId(2, idBox, action, sessionJpt, find);
            InputId(3, idBox, action, sessionJpt, find);
            InputId(4, idBox, action, sessionJpt, find);

            sessionJpt.FindElementByName("Mesh Size").Click();
            sessionJpt.FindElementByName("Keep Entity").Click();
            sessionJpt.FindElementByName("OK").Click();
            

            sessionJpt.FindElementByName("Meshing").Click();
            sessionJpt.FindElementByName("Surface Meshing").Click();
            typeInBox(1001, 15, sessionJpt);
            typeInBox(1002, 1, sessionJpt);
            typeInBox(1003, 30, sessionJpt);
            typeInBox(1004, 1, sessionJpt);
            
            sessionJpt.FindElementByName("OK").Click();
            /*
            System.Threading.Thread.Sleep(10000);

            WindowsElement completeDialog = sessionJpt.FindElementByName("Meshing Surf Meshing");
            action = new Actions(sessionJpt);
            action.MoveToElement(completeDialog);
            action.MoveToElement(completeDialog, completeDialog.Size.Width / 2 + 70,
                completeDialog.Size.Height / 2 + 60)
                .Click();
            action.Perform();
            */

            //sessionJpt.FindElementByName("OK").Click();
            System.Threading.Thread.Sleep(20000);
            var okBtns = sessionJpt.FindElementsByName("OK");
            if (okBtns[0].Location.Y < okBtns[1].Location.Y)
            {
                okBtns[0].Click();
            }
            else
            {
                okBtns[1].Click();
            }
            Debug.WriteLine(okBtns.Count);
            /*
            foreach (WindowsElement okBtn in okBtns)
            {
                Debug.WriteLine(okBtn.Location.Y);
            }
            */





        }
        public static void InputId(int faceId, WindowsElement idBox, Actions action,
                            WindowsDriver<WindowsElement> sessionJpt, WindowsElement find)
        {
            idBox.SendKeys(Keys.Control + "a" + Keys.Control);
            idBox.SendKeys(Convert.ToString(faceId));
            action = new Actions(sessionJpt);
            action.MoveToElement(find);
            action.MoveToElement(find, find.Size.Width / 2 - 30, find.Size.Height / 3 - 8).Click().Perform();
        }

        public static void typeInBox(int id, double keys, WindowsDriver<WindowsElement> sessionJpt)
        {
            var temp = sessionJpt.FindElementByAccessibilityId($"{id}");
            temp.SendKeys(Keys.Control + "a" + Keys.Control);
            temp.SendKeys($"{keys}");
        }
    }
}
