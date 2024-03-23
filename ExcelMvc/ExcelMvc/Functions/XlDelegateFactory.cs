using ExcelMvc.Interfaces;
using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    public static class XlDelegateFactory
    {
        public static IntPtr Make(MethodInfo method)
        {
            var count = method.GetParameters().Length;
            var type = method.ReturnType.Equals(typeof(void)) ?
                XlDelegateTypes.XlActions[count] : XlDelegateTypes.XlFunctions[count];

            var convert = typeof(Converter).GetMethod("IntPtr2Double");

            var outerParameters = method.GetParameters().Select(x => Expression.Parameter(typeof(IntPtr), x.Name));
            var innerParameters = outerParameters.Select(x=> convert == null ? x : (Expression) Expression.Call(convert, x)).ToArray(); 

            var p1 = Expression.Parameter(typeof(double), "a");
            var p2 = Expression.Parameter(typeof(double), "b");
            var p3 = Expression.Call(method, new[] { p1, p2 });
            var e = Expression.Lambda(type, p3, p1, p2).Compile();
            return Marshal.GetFunctionPointerForDelegate(e);
        }
    }
}
