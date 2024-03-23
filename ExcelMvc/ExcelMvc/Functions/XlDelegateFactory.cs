using ExcelMvc.Interfaces;
using System;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public static class XlDelegateFactory
    {
        public static IntPtr Make(MethodInfo method, FunctionAttribute function)
        {
            var count = method.GetParameters().Length;
            var type = method.ReturnType.Equals(typeof(void)) ?
                XlDelegateTypes.XlActions[count] : XlDelegateTypes.XlFunctions[count];

            var p1 = Expression.Parameter(typeof(double), "a");
            var p2 = Expression.Parameter(typeof(double), "b");
            var p3 = Expression.Call(method, new[] { p1, p2 });
            var e = Expression.Lambda(type, p3, p1, p2).Compile();
            return Marshal.GetFunctionPointerForDelegate(e);
        }
    }
}
