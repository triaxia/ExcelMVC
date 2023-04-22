namespace ExcelMvc.Rtd
{
    public delegate void UpdateNotify();

    /// <summary>
    /// Defines the implementation interface of a real-time data (RTD) server.
    /// </summary>
    public interface IRtdServerImpl
    {
        /// <summary>
        /// Called immediately after a RTD server is instantiated. 
        /// </summary>
        /// <param name="updateNotify">The method to be called whenever topic values are updated.</param>
        /// <returns>A negative value or zero indicates failure to start the server, a positive value
        /// indicates success.</returns>
        int Start(UpdateNotify updateNotify);

        /// <summary>
        /// Adds a new topic to the RTD server.
        /// </summary>
        /// <param name="topicId"></param>
        /// <param name="args"></param>
        /// <returns>The topic value</returns>
        object Connect(int topicId, string[] args);

        /// <summary>
        /// Notifies the RTD server that a topic is no longer in use.
        /// </summary>
        /// <param name="topicId"></param>
        void Disconnect(int topicId);

        /// <summary>
        /// Gets the updated topic values in the RTD server
        /// </summary>
        /// <returns>A two dimentional array with the first demension being the topic ids and 
        /// the second demension being the corresponding topic values.</returns>
        object[,] GetTopicValues();

        /// <summary>
        /// Indicates if the real-time data server (RTD) is still active.
        /// </summary>
        /// <returns>Zero or a negative number indicates failure; a positive number indicates that the server is active</returns>
        int Heartbeat();

        /// <summary>
        /// Terminates the real-time data (RTD) server.
        /// </summary>
        void Terminate();
    }
}
