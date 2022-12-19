using TestExcel.Models.Views;

namespace TestExcel.InterfacesServices
{
    /// <summary>
    /// IExcel core interface to read and write to excel files
    /// This interface is application specific
    /// uses data models, formats and styles specific to this app
    /// </summary>
    public interface IExcel
    {
        /// <summary>
        /// Read read spreadsheet
        /// </summary>
        /// <param name="fnap">filename and path</param>
        /// <returns>Tuple of message and list of data rows</returns>
        Tuple<string, List<TestView>> Read(string fnap);
        /// <summary>
        /// Write write spreadsheet
        /// </summary>
        /// <param name="dataRows">list of data rows</param>
        /// <param name="fnap">file name and path</param>
        /// <returns>message</returns>
        string Write(List<TestView> dataRows, string fnap);
    }
}