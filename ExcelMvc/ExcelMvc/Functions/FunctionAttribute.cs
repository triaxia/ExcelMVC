using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMvc.Functions
{
	/// <summary>
	/// Decorates functions that are to be exported as User Defined Functions.
	/// https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
	/// </summary>
	[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public class FunctionAttribute : Attribute
	{
		/// <summary>
		/// Specifies which category that the function should be listed in the function wizard.
		/// </summary>
		/// 
		public string Category;
		/// <summary>
		/// The function name as it will appear in the Function Wizard.
		/// </summary>
		public string Name { get; set; }
		/// <summary>
		/// The Description of the function when it is selected in the Function Wizard.
		/// </summary>
		public string Description { get; set; } = "";
		/// <summary>
		/// The help infomation displayed when the Help button is clicked.
		/// It can be in either "chm-file!HelpContextID" or "https://address/path_to_file_in_site!0". 
		/// </summary>
		public string HelpTopic { get; set; } = "";

		/// <summary>
		/// Indicates the type of function, 0, 1 or 2.
		/// </summary>
		public int FunctionType { get; set; } = 1;

		/// <summary>
		/// Registers the function as volatile, i.e. recalculates every time the worksheet recalculates.
		/// (pxTypeText +='!')
		/// </summary>
		public bool IsVolatile { get; set; } = false;

		/// <summary>
		/// Registers the function as macro sheet equivalent, handling uncalculated cells.
		/// pxTypeText +='#'
		/// </summary>
		public bool IsMacro { get; set; } = false;

		/// <summary>
		/// Registers the function as an asynchronous function.
		/// (pxTypeText=>(pxArgsTypeText)X)
		/// </summary>
		public bool IsAnyc { get; set; } = false;

		/// <summary>
		/// Indiates the function is thread-safe.
		/// (pxTypeTex +='$')
		/// </summary>
		public bool IsThreadSafe { get; set; } = false;

		/// <summary>
		/// Indiates the function is cluster-safe.
		/// (pxTypeText += '&')
		/// </summary>
		public bool IsClusterSafe { get; set; } = false;

		public FunctionAttribute()
		{
		}
	}
}
