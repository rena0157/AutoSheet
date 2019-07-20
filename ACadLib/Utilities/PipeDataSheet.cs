using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ACadLib.Utilities
{
    /// <summary>
    /// Pipe Data sheet Class that represents a pipe data sheet in Excel
    /// </summary>
    public class PipeDataSheet : DesignSheet
    {
        #region Excel Ranges

        /// <summary>
        /// The Range that represents the Whole PipeData Sheet
        /// </summary>
        public Range PipeDataRange { get; }

        /// <summary>
        /// The Range that represents the From Column
        /// </summary>
        public Range FromRange { get; }

        /// <summary>
        /// The Range that represents the To Column
        /// </summary>
        public Range ToRange { get; }

        /// <summary>
        /// The Range that represents the FromTo Column
        /// </summary>
        public Range FromToRange { get; }

        /// <summary>
        /// The Range that represents the Start Invert Column
        /// </summary>
        public Range StartInvRange { get; }

        public Range EndInvRange { get; }

        /// <summary>
        /// The Range that Represents the Slope Column
        /// </summary>
        public Range SlopeRange { get; }

        /// <summary>
        /// The Range that Represents the Length Column
        /// </summary>
        public Range LengthRange { get; }

        /// <summary>
        /// The Range that represents the Inner Diameter Column
        /// </summary>
        public Range InnerDiameterRange { get; }

        /// <summary>
        /// The Range that Represents the Handle Column
        /// </summary>
        public Range HandleRange { get; }

        private const string PipeDataName = "PipeData";
        private const string FromName = "PipeData_From";
        private const string ToName = "PipeData_To";
        private const string FromToName = "PipeData_FromToId";
        private const string StartInvName = "PipeData_StartInv";
        private const string EndInvName = "PipeData_EndInv";
        private const string SlopeName = "PipeData_Slope";
        private const string LengthName = "PipeData_Length";
        private const string InnerDiameterName = "PipeData_InnerDiameter";
        private const string HandleName = "PipeData_Handle";

        #endregion


        #region Public Members

        /// <summary>
        /// The name of the pipe Data sheet
        /// </summary>
        public const string PipeDataSheetName = "PipeData";

        #endregion

        /// <summary>
        /// Get a range from a column in the PipeData Sheet
        /// </summary>
        /// <param name="columnNumber">The column Number (A = 1)</param>
        /// <param name="upperBound">The upper bound of the Rows (Inclusive)</param>
        /// <param name="lowerBound">The lower bound of the Rows (Inclusive)</param>
        /// <returns>Returns: A range that represents a sub range of the column number that was passed</returns>
        public Range GetRangeFromColumn(int columnNumber, int upperBound, int lowerBound)
        {
            var columnA1 = (char) ( columnNumber - 1 + 'A' );
            return PipeDataRange.Range[$"{columnA1}{upperBound}:{columnA1}{lowerBound}"];
        }

        #region Constructors

        /// <summary>
        /// Default Constructor that loads the workbook and pipe data sheets
        /// </summary>
        /// <param name="filename">The filename of the excel file</param>
        /// <param name="pipeDataSheetName">The sheet name of the pipe data sheet</param>
        public PipeDataSheet(string filename, string pipeDataSheetName) : base(filename, pipeDataSheetName)
        {
            try
            {
                PipeDataRange = XlWorkbookNames[PipeDataName];
                FromRange = XlWorkbookNames[FromName];
                ToRange = XlWorkbookNames[ToName];
                StartInvRange = XlWorkbookNames[StartInvName];
                EndInvRange = XlWorkbookNames[EndInvName];
                SlopeRange = XlWorkbookNames[SlopeName];
                LengthRange = XlWorkbookNames[LengthName];
                InnerDiameterRange = XlWorkbookNames[InnerDiameterName];
                HandleRange = XlWorkbookNames[HandleName];
            }
            catch ( KeyNotFoundException )
            {

            }
        }

        #endregion

    }
}
