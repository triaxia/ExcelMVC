﻿using ExcelMvc.Interfaces;
using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static class DelegateFactory
    {
        public static Delegate MakeOuterDelegate(MethodInfo method)
        {
            var instance = new LazyDelegate(() => MakeInnerDelegate(method));
            var count = method.GetParameters().Length;
            if (method.ReturnType.Equals(typeof(void)))
            {
                return Delegate.CreateDelegate(ActionDelegate.Actions[count],
                    instance, LazyDelegate.GetActionMethod(count));
            }
            else
            {
                return Delegate.CreateDelegate(FunctionDelegate.Functions[count],
                    instance, LazyDelegate.GetFunctionMethod(count));
            }
        }

        public static Delegate MakeInnerDelegate(MethodInfo method)
        {
            var expressions = method.GetParameters()
                .Select(x => (x.ParameterType, expression: Expression.Parameter(typeof(IntPtr), x.Name)));
            
            var outerParameters = expressions.Select(x => x.expression)
                .ToArray();
            var innerParameters = expressions.Select(x => Expression.Call(XlMarshalContext.IntPtr2Parameter(x.ParameterType), x.expression))
                .ToArray();
            var innerCall = Expression.Call(method, innerParameters);

            var ex = Expression.Variable(typeof(Exception), "ex");
            var exHandler = Expression.Call(XlMarshalException.HandlerMethod, ex);

            if (method.ReturnType == typeof(void))
            {
                var catcher = Expression.Block(exHandler, Expression.Empty()); // TODO
                var body = Expression.TryCatch(innerCall, Expression.Catch(ex, catcher));
                var delegateType = ActionDelegate.Actions[outerParameters.Length];
                return Expression.Lambda(delegateType, body, method.Name, outerParameters).Compile();
            }
            else
            {
                var context = Expression.Variable(typeof(XlMarshalContext), "context");
                var value = Expression.Call(typeof(XlMarshalContext), nameof(XlMarshalContext.GetThreadInstance), null);
                innerCall = Expression.Call(context, XlMarshalContext.Result2IntPtr(method.ReturnType), innerCall);
                var catcher = Expression.Block(exHandler, Expression.Constant(IntPtr.Zero)); // TODO

                var delegateType = FunctionDelegate.Functions[outerParameters.Length];
                var body = Expression.Block(
                    typeof(IntPtr),
                    new ParameterExpression[] { context },
                    Expression.Assign(context, value),
                    Expression.TryCatch(innerCall, Expression.Catch(ex, catcher)));
                return Expression.Lambda(delegateType, body, method.Name, outerParameters).Compile();
            }
        }
    }
}