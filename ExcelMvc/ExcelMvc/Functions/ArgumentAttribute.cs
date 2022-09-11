using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMvc.Functions
{
	/// <summary>
	/// Decorates arguments of exported functions.
	/// </summary>
	[AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public class ExcelArgumentAttribute : Attribute
	{
		public string Name { get; set; }
		public string Description { get; set; }
        
		public ExcelArgumentAttribute()
		{
		}
		public void Test()
        {

        }
	}
}
