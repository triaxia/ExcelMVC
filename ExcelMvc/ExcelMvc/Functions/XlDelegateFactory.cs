using ExcelMvc.Interfaces;
using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static class XlDelegateFactory
    {
        public static Delegate MakeOuterDelegate(MethodInfo method, FunctionAttribute function)
        {
            var inner = new DelegateInvoke(() => MakeInnerDelegate(method, function));
            var count = method.GetParameters().Length;  
            if (method.ReturnType.Equals(typeof(void)))
            {
                return Delegate.CreateDelegate(XlActionDelegate.XlActions[count],
                    inner, typeof(DelegateInvoke).GetMethod($"Action{count}"));
            }
            else
            {
                return Delegate.CreateDelegate(XlFunctionDelegate.XlFunctions[count],
                    inner, typeof(DelegateInvoke).GetMethod($"Function{count}"));
            }
        }

        public static Delegate MakeInnerDelegate(MethodInfo method, FunctionAttribute function)
        {
            // a lot to do here...
            var count = method.GetParameters().Length;
            var type = method.ReturnType.Equals(typeof(void)) ?
                XlActionDelegate.XlActions[count] : XlFunctionDelegate.XlFunctions[count];

            var convert = typeof(Converter).GetMethod("IntPtr2Double");

            var outerParameters = method.GetParameters().Select(x => Expression.Parameter(typeof(IntPtr), x.Name));
            var innerParameters = outerParameters.Select(x => convert == null ? x : (Expression)Expression.Call(convert, x)).ToArray();

            var p1 = Expression.Parameter(typeof(double), "a");
            var p2 = Expression.Parameter(typeof(double), "b");
            var p3 = Expression.Call(method, new[] { p1, p2 });
            var e = Expression.Lambda(type, p3, p1, p2).Compile();
            return e;
            //return Marshal.GetFunctionPointerForDelegate(e);
        }
    }
}
