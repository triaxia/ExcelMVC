using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace Mvc
{
    /// <summary>
    /// Decorates functions that are to be exported as User Defined Functions.
    /// https://docs.microsoft.com/en-us/office/client-developer/excel/xlfregister-form-1
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public class FunctionAttribute : Attribute
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
        public byte FunctionType = 1;

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
        /// Indicates if the function is listed in the Function Wizard.
        /// </summary>
        public bool IsHidden = false;

        /// <summary>
        /// Registers the function as an asynchronous function.
        /// (pxTypeText=>(pxArgsTypeText)X)
        /// </summary>
        public bool IsAsync;

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

}
