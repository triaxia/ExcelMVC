using System.Reflection;
using System.Runtime.InteropServices;

namespace Function.Interfaces
{
    /// <summary>
    /// Represents the properties of a function argument.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct ArgumentDefinition
    {
        /// <summary>
        /// The name of the argument.
        /// </summary>
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Name;

        /// <summary>
        /// The description of the argument.
        /// </summary>
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Description;

        /// <summary>
        /// The type of the argument
        /// </summary>
        [MarshalAs(UnmanagedType.LPWStr)]
        public string Type;

        /// <summary>
        /// Initialises a new instance of <see cref="ArgumentDefinition"/>
        /// </summary>
        /// <param name="parameter"></param>
        /// <param name="argument"></param>
        public ArgumentDefinition(ParameterInfo parameter, IArgumentAttribute argument)
        {
            if (argument == null)
            {
                Name = parameter.Name;
                Description = "";
            }
            else
            {
                Name = argument.Name ?? parameter.Name;
                Description = argument.Description ?? "";
            }
            Type = parameter.ParameterType.FullName;
        }
        /// <summary>
        /// Indicates if an argument is optional.
        /// </summary>
        public bool IsOptionalArg => Name.StartsWith("[") && Name.EndsWith("]");
    }
}
