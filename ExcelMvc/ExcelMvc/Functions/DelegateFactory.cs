using ExcelMvc.Interfaces;
using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static class DelegateFactory
    {
        public static Delegate MakeOuterDelegate(MethodInfo method, FunctionAttribute function)
        {
            var instance = new DelegateInvoke(() => MakeInnerDelegate(method, function));
            var count = method.GetParameters().Length;
            if (method.ReturnType.Equals(typeof(void)))
            {
                return Delegate.CreateDelegate(ActionDelegate.Actions[count],
                    instance, typeof(DelegateInvoke).GetMethod($"Action{count}"));
            }
            else
            {
                return Delegate.CreateDelegate(FunctionDelegate.Functions[count],
                    instance, typeof(DelegateInvoke).GetMethod($"Function{count}"));
            }
        }

        public static Delegate MakeInnerDelegate(MethodInfo method, FunctionAttribute function)
        {
            var convert = typeof(Converter).GetMethod("IntPtr2Double");

            var expressions = method.GetParameters().Select(x => (x.ParameterType, expression: Expression.Parameter(typeof(IntPtr), x.Name)));
            
            var outerParameters = expressions.Select(x => x.expression);
            var innerParameters = expressions.Select(x => Expression.Call(XlMarshalContext.IntPtr2ParameterMethod(x.ParameterType), x.expression));

            /*
            // a lot to do here...
            var count = method.GetParameters().Length;
            var type = method.ReturnType.Equals(typeof(void)) ?
                ActionDelegate.Actions[count] : FunctionDelegate.Functions[count];
            */

            //assign, catch ...
            //var e = Expression.Lambda(type, p3, p1, p2).Compile();
            return null;
        }
    }
}
