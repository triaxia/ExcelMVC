using System;

namespace Addin.Interfaces
{
	/// <summary>
	/// Decorates arguments of exported functions.
	/// </summary>
	[AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
	public class ArgumentAttribute : Attribute
	{
		public string Name;
		public string Description;
	}
}
