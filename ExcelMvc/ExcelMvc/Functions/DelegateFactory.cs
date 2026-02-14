/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; either version 2 of
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program;
if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor,
Boston, MA 02110-1301 USA.
*/
using Function.Interfaces;
using System;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMvc.Functions
{
    public static partial class DelegateFactory
    {
        public static Delegate MakeOuterDelegate(MethodInfo method, FunctionDefinition function)
        {
            var instance = new LazyDelegate(() =>
            {
                try
                {
                    return MakeInnerDelegate(method, function);
                }
                catch (Exception ex)
                {
                    XlMarshalExceptionHandler.HandleException(ex);
                    return MakeZeroDelegate(method);
                }
            });
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

        public static Delegate MakeInnerDelegate(MethodInfo method, FunctionDefinition function)
        {
            var parameters = method.GetParameters();
            var outerParameters = new ParameterExpression[parameters.Length];
            var variables = new ParameterExpression[parameters.Length];
            var varLines = new BinaryExpression[parameters.Length];
            for (var index = 0; index < parameters.Length; index++)
            {
                outerParameters[index] = Expression.Parameter(typeof(IntPtr), parameters[index].Name);
                var innerParameter = Expression.Call(XlMarshalContext.IncomingConverter(parameters[index].ParameterType)
                    , outerParameters[index]
                    , Expression.Constant(parameters[index])
                    , Expression.Constant(function.Arguments != null && function.Arguments[index].IsOptionalArg));
                var variable = Expression.Variable(parameters[index].ParameterType, $"_{parameters[index].Name}_");
                variables[index] = variable;
                varLines[index] = Expression.Assign(variable, innerParameter);
            }

            var methodCall = (Expression)Expression.Call(method, variables);

            var instanceProperty = Expression.Property(null, typeof(FunctionHost), nameof(FunctionHost.Instance));
            var executingEventRaised = Expression.Property(instanceProperty, nameof(IFunctionHost.ExecutingEventRaised));

            var args = variables.Select(x=> Expression.Convert(x, typeof(object))).ToArray();
            var stopwatch = Expression.Variable(typeof(Stopwatch), "stopwatch");
            var startNew = Expression.Call(typeof(Stopwatch), nameof(Stopwatch.StartNew), null);
            var assignStopWatch = Expression.Assign(stopwatch, startNew);
            var elapsedProperty = Expression.Property(stopwatch, nameof(Stopwatch.Elapsed));
            var logging = Expression.Call(LoggingMethod(args.Length)
                , new Expression[] { Expression.Constant(function.Name), Expression.Constant(method), elapsedProperty }.Concat(args));

            Expression innerCall;
            if (method.ReturnType == typeof(void))
            {
                var loggingPath = Expression.Block(typeof(void), new[] { stopwatch },
                    assignStopWatch, methodCall, logging);
                var conditional = Expression.Condition(executingEventRaised, loggingPath, methodCall);
                innerCall = Expression.Block(method.ReturnType, variables, varLines.Cast<Expression>().Concat(new[] {conditional}));
            }
            else
            {
                var result = Expression.Variable(method.ReturnType, "result");
                var assignResult = Expression.Assign(result, methodCall);
                var loggingPath = Expression.Block(method.ReturnType, new[] {stopwatch, result},
                    assignStopWatch, assignResult, logging, result);
                var conditional = Expression.Condition(executingEventRaised, loggingPath, methodCall);
                innerCall = Expression.Block(method.ReturnType, variables, varLines.Cast<Expression>().Concat(new[] { conditional }));
            }

            var ex = Expression.Variable(typeof(Exception), "ex");
            if (method.ReturnType == typeof(void))
            {
                var handler = Expression.Block(Expression.Call(XlMarshalExceptionHandler.HandlerMethod, ex)
                    , Expression.Empty());
                var body = Expression.TryCatch(innerCall, Expression.Catch(ex, handler));
                var delegateType = ActionDelegate.Actions[outerParameters.Length];
                return Expression.Lambda(delegateType, body, method.Name, outerParameters).Compile();
            }
            else
            {
                var context = Expression.Variable(typeof(XlMarshalContext), "context");
                var value = Expression.Call(typeof(XlMarshalContext), nameof(XlMarshalContext.GetThreadInstance), null);
                innerCall = Expression.Call(context, XlMarshalContext.OutgoingConverter(method.ReturnType), innerCall);
                var handler = Expression.Call(context, XlMarshalContext.ExceptionConverter(method.ReturnType)
                    , Expression.Call(XlMarshalExceptionHandler.HandlerMethod, ex));
                var body = Expression.Block(
                    typeof(IntPtr),
                    new ParameterExpression[] { context },
                    Expression.Assign(context, value),
                    Expression.TryCatch(innerCall, Expression.Catch(ex, handler)));

                var delegateType = FunctionDelegate.Functions[outerParameters.Length];
                var lambda = Expression.Lambda(delegateType, body, method.Name, outerParameters);
                return lambda.Compile();
            }
        }

        public static Delegate MakeZeroDelegate(MethodInfo method)
        {
            var outerParameters = method.GetParameters()
                .Select(x => Expression.Parameter(typeof(IntPtr), x.Name))
                .ToArray();

            if (method.ReturnType == typeof(void))
            {
                var delegateType = ActionDelegate.Actions[outerParameters.Length];
                return Expression.Lambda(delegateType, Expression.Empty(), method.Name, outerParameters).Compile();
            }
            else
            {
                var delegateType = FunctionDelegate.Functions[outerParameters.Length];
                var lambda = Expression.Lambda(delegateType, Expression.Constant(IntPtr.Zero), method.Name, outerParameters);
                return lambda.Compile();
            }
        }
    }
}
