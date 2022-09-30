using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    /// <summary>
    /// Decorates functions that are to be exported as User Defined Functions.
    /// https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
	public class ExcelFunctionAttribute : Attribute
	{
		/// <summary>
		/// Specifies which category that the function should be listed in the function wizard.
		/// </summary>
		public string Category;
		/// <summary>
		/// The function name as it will appear in the Function Wizard.
		/// </summary>
		public string Name;
		/// <summary>
		/// The Description of the function when it is selected in the Function Wizard.
		/// </summary>
		public string Description;
		/// <summary>
		/// The help infomation displayed when the Help button is clicked.
		/// It can be in either "chm-file!HelpContextID" or "https://address/path_to_file_in_site!0". 
		/// </summary>
		public string HelpTopic;

		/// <summary>
		/// Indicates the type of function, 0, 1 or 2.
		/// </summary>
		public int FunctionType = 1;

		/// <summary>
		/// Registers the function as volatile, i.e. recalculates every time the worksheet recalculates.
		/// (pxTypeText +='!')
		/// </summary>
		public bool IsVolatile;

		/// <summary>
		/// Registers the function as macro sheet equivalent, handling uncalculated cells.
		/// pxTypeText +='#'
		/// </summary>
		public bool IsMacro;

		/// <summary>
		/// Registers the function as an asynchronous function.
		/// (pxTypeText=>(pxArgsTypeText)X)
		/// </summary>
		public bool IsAnyc;

		/// <summary>
		/// Indiates the function is thread-safe.
		/// (pxTypeTex +='$')
		/// </summary>
		public bool IsThreadSafe;

		/// <summary>
		/// Indiates the function is cluster-safe.
		/// (pxTypeText += '&')
		/// </summary>
		public bool IsClusterSafe;
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public struct ExcelFunction
	{
		public string Category;
		public string Name;
		public string Description;
		public string HelpTopic;
		public int FunctionType;
		public bool IsVolatile;
		public bool IsMacro;
		public bool IsAnyc;
		public bool IsThreadSafe;
		public bool IsClusterSafe;
		public int ArgumentCount;
		public ExcelArgument[] Arguments;

		public ExcelFunction(ExcelFunctionAttribute rhs,
			ExcelArgument[] arguments)
		{
			Category = rhs.Category;
			Name = rhs.Name;
			Description = rhs.Description;
			HelpTopic = rhs.HelpTopic;
			FunctionType = rhs.FunctionType;
			IsVolatile = rhs.IsVolatile;
			IsMacro = rhs.IsMacro; 
			IsAnyc = rhs.IsAnyc; 
			IsThreadSafe = rhs.IsThreadSafe; 
			IsClusterSafe = rhs.IsClusterSafe;
			Arguments = arguments;
			ArgumentCount = arguments?.Length ?? 0;
		}
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct ExcelFunctions
	{
		public int Count;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 2000)]
		public ExcelFunction[] Functions;
	}
}
