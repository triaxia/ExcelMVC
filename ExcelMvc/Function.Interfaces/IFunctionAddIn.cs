namespace Function.Interfaces
{
    /// <summary>
    /// Defines the behaviour of function add-in components;
    /// </summary>
    public interface IFunctionAddIn
    {
        /// <summary>
        /// Opens the add-in.
        /// </summary>
        void Open();

        /// <summary>
        /// Closes the add-in.
        /// </summary>
        void Close();

        /// <summary>
        /// Defines the oder of <see cref="Open"/>. Add-ins with larger rankings are 
        /// opened before those with smaller rankings. The <see cref="Ranking"/>
        /// is typicallly used to open the root add-in in front of others, as it is
        /// usually the one who is responsible for initialising configuraitons and 
        /// dependencies etc.
        /// </summary>
        int Ranking { get; }
    }
}
