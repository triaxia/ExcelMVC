using System;
using System.Runtime.InteropServices;

namespace ExcelMvc.Functions
{
    /// <summary>
    /// Decorates arguments of exported functions.
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public class ExcelArgumentAttribute : Attribute
	{
		public string Name;
		public string Description = "";
        
		public ExcelArgumentAttribute()
		{
		}
	}
}
