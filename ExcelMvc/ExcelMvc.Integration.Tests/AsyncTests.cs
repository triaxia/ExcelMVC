﻿using Function.Interfaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelMvc.Integration.Tests
{
    [TestClass]
    public class AsyncTests
    {
        private static object PostActionValue = null;
        [Function()]
        public static object uPostAction([Argument(Name = "[v1]")] object v1)
        {
            FunctionHost.Instance.PostAction(state =>
            {
                PostActionValue = state;
            }, v1);

            return PostActionValue;
        }

        [TestMethod]
        public void uPostAction()
        {
            using (var excel = new ExcelLoader())
            {
                var value = Guid.NewGuid().ToString();
                excel.Application.Run("uPostAction", value);
                Thread.Sleep(1000);
                var result = (string)excel.Application.Run("uPostAction");
                Assert.AreEqual(value, result);
            }
        }

        private static object PostMacroValue = null;
        [Function()]
        public static object uPostMacro([Argument(Name = "[v1]")] object v1)
        {
            FunctionHost.Instance.PostMacro(state =>
            {
                PostMacroValue = state;
            }, v1);

            return PostMacroValue;
        }

        [TestMethod]
        public void uPostMacro()
        {
            using (var excel = new ExcelLoader())
            {
                var value = Guid.NewGuid().ToString();
                excel.Application.Run("uPostMacro", value);
                Thread.Sleep(1000);
                var result = (string)excel.Application.Run("uPostMacro");
                Assert.AreEqual(value, result);
            }
        }

        [Function(IsAsync = true)]
        public static void uAsyncFunction([Argument(Name = "[v1]")] object v1, IntPtr handle)
        {
            var h = FunctionHost.Instance.GetAsyncHandle(handle);
            Task.Factory.StartNew(state =>
            {
                var args = (object[])state;
                FunctionHost.Instance.SetAsyncValue((IntPtr)args[0], args[1]);
            }, new object[] { h, v1 });
        }

        [TestMethod]
        public void uAsyncFunction()
        {
            using (var excel = new ExcelLoader())
            {
                var value = Guid.NewGuid().ToString();
                var result = excel.Application.Run("uAsyncFunction", value);
                Assert.AreEqual(value, result);
            }
        }
    }
}