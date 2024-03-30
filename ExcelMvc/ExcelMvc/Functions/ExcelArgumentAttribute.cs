using System;

namespace ExcelMvc.Functions
{
	/// <summary>
	/// Decorates arguments of exported functions.
	/// </summary>
	[AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
	public class ExcelArgumentAttribute : Attribute
	{
		public string Name;
		public string Description;
	}
}
