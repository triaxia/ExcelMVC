using Function.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMvc.Functions
{
    public static class XLRegistration
	{
		private static readonly object[] ExcelMvcAttach =
			{ "ExcelMvcAttach", "I", "ExcelMvcAttach", "", 1, "ExcelMvc", "", "", "Attach Excel to ExcelMvc", "", "" };
		private static readonly object[] ExcelMvcDetach =
			{ "ExcelMvcDetach", "I", "ExcelMvcDetach", "", 1, "ExcelMvc", "", "", "Detach Excel from ExcelMvc", "", "" };
		private static readonly object[] ExcelMvcShow =
			{ "ExcelMvcShow", "I", "ExcelMvcShow", "", 1, "ExcelMvc", "", "", "Shows the ExcelMvc window", "", "" };
		private static readonly object[] ExcelMvcHide =
			{ "ExcelMvcHide", "I", "ExcelMvcHide", "", 1, "ExcelMvc", "", "", "Hides the ExcelMvc window", "", "" };
		private static readonly object[] ExcelMvcClick =
			{ "ExcelMvcClick", "I", "ExcelMvcRunCommandAction", "", 2, "ExcelMvc", "", "", "Called by a command", "", "" };
		private static readonly object[] ExcelMvcRun =
			{ "ExcelMvcRun", "I", "ExcelMvcRun", "", 2, "ExcelMvc", "", "", "Runs the next action in the async queue", "", "" };

		private static readonly List<object> FunctionIds = new List<object>();

		public static void UnregisterAll()
		{
			foreach (var id in FunctionIds)
			{
				(var status, var _ )  = XLCall.Call(XLFunctions.xlfUnregister, new[] { id });
				XLCall.EnsureSuccessStatusCode(status, () => $"Unregister function failed with status {status}.");

			}
			FunctionIds.Clear();
		}

		public static void Register(FunctionDefinitions functions)
		{
			(var status, var xll) = XLCall.Call(XLFunctions.xlGetName, System.Array.Empty<object>());
			XLCall.EnsureSuccessStatusCode(status, () => $"{nameof(XLFunctions.xlGetName)} failed with status code {status}.");

			Register(nameof(ExcelMvcAttach), new[] { xll }.Concat(ExcelMvcAttach).ToArray());
            Register(nameof(ExcelMvcDetach), new[] { xll }.Concat(ExcelMvcDetach).ToArray());
            Register(nameof(ExcelMvcShow), new[] { xll }.Concat(ExcelMvcShow).ToArray());
            Register(nameof(ExcelMvcHide), new[] { xll }.Concat(ExcelMvcHide).ToArray());
            Register(nameof(ExcelMvcClick), new[] { xll }.Concat(ExcelMvcClick).ToArray());
            Register(nameof(ExcelMvcRun), new[] { xll }.Concat(ExcelMvcRun).ToArray());
			for (var i = 0; i < functions.FunctionCount; i++)
				Register(functions.Items[i], i, xll);
        }

        public static object Register(string name, object parameters)
		{
			(var status, var result) = XLCall.Call(XLFunctions.xlfRegister, parameters);
			XLCall.EnsureSuccessStatusCode(status, () => $"Registering \"{name}\" failed with status {status}.");
			FunctionIds.Add(result);
			return result;
		}

		public static object Register(FunctionDefinition function, int index, object xll)
		{
			return null;
		}
    }
}
