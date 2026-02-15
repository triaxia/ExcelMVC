using Function.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMvc.Functions
{
    public static class XLRegistration
    {
        private static readonly object[] ExcelMvcAttach =
            { "ExcelMvcAttach", "I", "ExcelMvcAttach", "", 1, "ExcelMvc", "", "", "Attach Excel to ExcelMvc", "" };
        private static readonly object[] ExcelMvcDetach =
            { "ExcelMvcDetach", "I", "ExcelMvcDetach", "", 1, "ExcelMvc", "", "", "Detach Excel from ExcelMvc", "" };
        private static readonly object[] ExcelMvcShow =
            { "ExcelMvcShow", "I", "ExcelMvcShow", "", 1, "ExcelMvc", "", "", "Shows the ExcelMvc window", "" };
        private static readonly object[] ExcelMvcHide =
            { "ExcelMvcHide", "I", "ExcelMvcHide", "", 1, "ExcelMvc", "", "", "Hides the ExcelMvc window", "" };
        private static readonly object[] ExcelMvcClick =
            { "ExcelMvcClick", "I", "ExcelMvcRunCommandAction", "", 2, "ExcelMvc", "", "", "Called by a command", "" };
        private static readonly object[] ExcelMvcRun =
            { "ExcelMvcRun", "I", "ExcelMvcRun", "", 2, "ExcelMvc", "", "", "Runs the next action in the async queue", "" };

        private static readonly List<object> FunctionIds = new List<object>();

        public static void UnregisterAll()
        {
            foreach (var id in FunctionIds)
            {
                (var status, var _) = XLCall.Call(XLFunctions.xlfUnregister, new[] { id });
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

        public static object Register(string name, object[] parameters)
        {
            (var status, var result) = XLCall.Call(XLFunctions.xlfRegister, parameters);
            XLCall.EnsureSuccessStatusCode(status, () => $"Registering \"{name}\" failed with status {status}.");
            FunctionIds.Add(result);
            return result;
        }

        public static object Register(FunctionDefinition function, int index, object xll)
        {
            /*
            https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
            LPXLOPER12 pxProcedure
            LPXLOPER12 pxTypeText
            LPXLOPER12 pxFunctionText
            LPXLOPER12 pxArgumentText
            LPXLOPER12 pxMacroType,
            LPXLOPER12 pxCategory
            LPXLOPER12 pxShortcutText
            LPXLOPER12 pxHelpTopic
            LPXLOPER12 pxFunctionHelp
            LPXLOPER12 pxArgumentHelp1
            LPXLOPER12 pxArgumentHelp2
            .
            LPXLOPER12 pxArgumentHelp255
            */
            var pxProcedure = $"udf{index}";

            (var pxArgumentText, var pxTypeText) = MakeArgumentList(function);

            var pxFunctionText = function.Name;
            var pxMacroType = function.IsHidden ? 0 : 1;

            var pxCategory = function.Category ?? "";
            var pxShortcutText = "";

            var pxHelpTopic = NormaliseHelpTopic(function);

            var count = 10 + function.ArgumentCount;

            var parameters = new object[count];

            parameters[0] = xll;
            parameters[1] = pxProcedure;
            parameters[2] = pxTypeText;
            parameters[3] = pxFunctionText;
            parameters[4] = pxArgumentText;
            parameters[5] = pxMacroType;
            parameters[6] = pxCategory;
            parameters[7] = pxShortcutText;
            parameters[8] = pxHelpTopic;

            const int DescriptionLimit = 252;
            parameters[9] = TruncateSentence(function.Description, DescriptionLimit);
            for (var idx = 0; idx < function.ArgumentCount; idx++)
                parameters[10 + idx] = TruncateSentence(function.Arguments[idx].Description, DescriptionLimit);
            return Register(function.Name, parameters);
        }

        private static (string names, string types) MakeArgumentList(FunctionDefinition function)
        {
            var types = function.IsAsync ? ">" : MakeTypeString(function.ReturnType, function.Name);
            var names = "";
            for (var idx = 0; idx < function.ArgumentCount; idx++)
            {
                if (idx > 0) names += ",";
                names += function.Arguments[idx].Name ?? "";
                types += MakeTypeString(function.Arguments[idx].Type, function.Arguments[idx].Name);
            }
            if (function.IsVolatile) types += "!";
            if (function.IsThreadSafe && !function.IsMacroType) types += "$";
            if (function.IsClusterSafe && !function.IsAsync) types += "&";
            if (function.IsMacroType) types += "#";

            return (names, types);
        }

        private static string MakeTypeString(string type, string argName)
        {
            bool Equals(string lhs, string rhs) => lhs.CompareTo(rhs) == 0;
            if (Equals(type, "System.Double")
                || Equals(type, "System.Float")
                || Equals(type, "System.UInt32")
                || Equals(type, "System.DateTime"))
                return "E";
            if (Equals(type, "System.Boolean"))
                return "L";
            if (Equals(type, "System.Int16")
                || Equals(type, "System.Byte")
                || Equals(type, "System.SByte"))
                return "M";
            if (Equals(type, "System.Int32")
                || Equals(type, "System.UInt16"))
                return "N";
            if (Equals(type, "System.String"))
                return "C%";
            if (Equals(type, "System.Double[,]")
                || Equals(type, "System.Double[]")
                || Equals(type, "System.DateTime[,]")
                || Equals(type, "System.DateTime[]")
                || Equals(type, "System.Int32[,]")
                || Equals(type, "System.Int32[]"))
                return "K%";
            if (Equals(type, "System.IntPtr"))
                return "X";
            return "Q";
        }

        private static string NormaliseHelpTopic(FunctionDefinition function)
        {
            var topic = function.HelpTopic ?? "";
            if (topic.Contains('!')) return topic;

            topic = topic.ToLower();

            if (topic.Contains("http://") || topic.Contains("https://"))
                topic += "!0";

            return topic;
        }

        private static string TruncateSentence(string text, int limit)
        {
            text = text ?? "";
            if (text.Length <= limit)
                return text;

            var words = text.Split(' ');
            var result = text;
            for (var i = words.Length - 1; i >= 0 && result.Length > limit; i--)
                result = string.Join(" ", words, 0, i);
            return $"{result}...";
        }
    }
}
