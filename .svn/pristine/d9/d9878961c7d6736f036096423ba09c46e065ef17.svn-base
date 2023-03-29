using System;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    /// <summary>
    /// An error or warning message from the server. The
    /// message may be tied to a specific test variable or variable value
    /// </summary>
    public class PDCMessage
    {
        private string myParameterName;
        private int myParameterNo;
        private string myVariableName;
        private int? myVariableNo;
        private int? myPosition;
        private int myExperimentIndex;

        private string myMessage;
        /*
         * The message type/loglevel used in the database is now an integer. 
         * But since the webservice interface still uses string we also have to use
         * a string.
         */
        private string myMessageType;

        /// <summary>
        /// The message code as it is returned by the server
        /// </summary>
        private string myMessageCode;
        /// <summary>
        /// string constant for PDC message type 'Warning'
        /// </summary>
        public const string TYPE_WARN = "3";
        /// <summary>
        /// string constant for PDC message type 'Error'
        /// </summary>
        public const string TYPE_ERROR = "4";
        /// <summary>
        /// string constant for PDC message type 'Debug'
        /// </summary>
        public const string TYPE_INFO_DEBUG = "1";
        /// <summary>
        /// string constant for PDC message type 'Fatal'
        /// </summary>
        public const string TYPE_FATAL = "5";
        /// <summary>
        /// string constant for PDC message type 'Fine'
        /// </summary>
        public const string TYPE_FINE = "2";

        #region constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public PDCMessage()
        {
        }

        /// <summary>
        /// Creates a new PDC message with the specified text of the specified type
        /// </summary>
        /// <param name="aMessage"></param>
        /// <param name="aMessageType"></param>
        public PDCMessage(string aMessage, string aMessageType)
        {
            myMessage = aMessage;
            myMessageType = aMessageType;
        }

        #endregion

        #region methods

        #region GetType
        /**
     * hack needed because WSDL does not define message type 
     * 
     */
        internal static string GetType(string msgid)
        {
            string ret = "";
            switch (msgid)
            {
                case PDCConstants.C_MSG_BAYMISSING:
                    ret = TYPE_INFO_DEBUG;
                    break;
                case PDCConstants.C_MSG_PDC_COMPOUND_DEFERRED:
                    ret = TYPE_INFO_DEBUG;
                    break;
                case PDCConstants.C_MSG_PDC_PREPARATION_DEFERRED:
                    ret = TYPE_INFO_DEBUG;
                    break;
                case PDCConstants.C_MSG_INVALID_FORMAT:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_MISMATCH_CP:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_MISMATCH_PM:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_PDC_INVALIDP_VALIDM:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_PDC_INVALIDC_VALIDP:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_PDC_PREPARATION_COP:
                    ret = TYPE_INFO_DEBUG;
                    break;
                case PDCConstants.C_MSG_MISSING_CNO:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_MISSING_PARAM:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_MISSING_PREPNO:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_MULTIPLE_PREPNO:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_UNKNOWN_COMPOUNDNO:
                    ret = TYPE_WARN;
                    break;
                case PDCConstants.C_MSG_UNKNOWN_MCNO:
                    ret = TYPE_ERROR;
                    break;
                case PDCConstants.C_MSG_UNKNOWN_PREPNO:
                    ret = TYPE_WARN;
                    break;
                default:
                    ret = null;
                    break;
            }
            return ret;
        }
        #endregion

        #endregion

        #region properties

        #region ExperimentIndex
        /// <summary>
        /// Associates the message with an experiment
        /// </summary>
        public int ExperimentIndex
        {
            get
            {
                return myExperimentIndex;
            }
            set
            {
                myExperimentIndex = value;
            }
        }
        #endregion

        #region LogLevel
        /// <summary>
        /// Returns the loglevel of the message
        /// </summary>
        public int LogLevel
        {
            get
            {
                if (myMessageType == null)
                {
                    return 0;
                }
                try
                {
                    return int.Parse(myMessageType);
                }
#pragma warning disable 0168
                catch (Exception e)
                {
                    return 0;
                }
#pragma warning restore 0168
            }
        }
        #endregion

        #region Message
        /// <summary>
        /// Property for the message string
        /// </summary>
        public string Message
        {
            get
            {
                return myMessage;
            }
            set
            {
                myMessage = value;
            }
        }
        #endregion

        #region MessageType
        /// <summary>
        /// Property for the message type
        /// </summary>
        public string MessageType
        {
            get
            {
                return myMessageType;
            }
            set
            {
                myMessageType = value;
            }
        }
        #endregion

        #region MessageTypeText
        /// <summary>
        /// Returns a readable text for the associated message type
        /// </summary>
        public string MessageTypeText
        {
            get
            {
                if (MessageType == null)
                {
                    return null;
                }
                switch (MessageType)
                {
                    case TYPE_ERROR: return "Error";
                    case TYPE_FATAL: return "Fatal";
                    case TYPE_WARN: return "Warning";
                    case TYPE_INFO_DEBUG: return "Info/Debug";
                    case TYPE_FINE: return "Fine";
                    default: return "Unspecified";
                }
            }
        }
        #endregion

        #region ParameterName
        /// <summary>
        /// Display name of the associated pdc parameter
        /// </summary>
        public string ParameterName
        {
            get
            {
                return myParameterName;
            }
            set
            {
                myParameterName = value;
            }
        }
        #endregion

        #region ParameterNo
        /// <summary>
        /// The identifier of the parameter. May be null if the PDCMessage is not directly bound
        /// to a specific parameter
        /// </summary>
        public int ParameterNo
        {
            get
            {
                return myParameterNo;
            }
            set
            {
                myParameterNo = value;
            }
        }
        #endregion

        #region MessageCode
        public string MessageCode
        {
            get { return myMessageCode; }
            set { myMessageCode = value; }
        }
        #endregion

        #region Position
        /// <summary>
        /// The position argument of the associated parameter if applicable
        /// </summary>
        public int? Position
        {
            get
            {
                return myPosition;
            }
            set
            {
                myPosition = value;
            }
        }
        #endregion

        #region VariableName
        /// <summary>
        /// The name of the associated test variable if applicable
        /// </summary>
        public string VariableName
        {
            get
            {
                return myVariableName;
            }
            set
            {
                myVariableName = value;
            }
        }
        #endregion

        #region VariableNo
        /// <summary>
        /// The no of the associated test variable if applicable
        /// </summary>
        public int? VariableNo
        {
            get
            {
                return myVariableNo;
            }
            set
            {
                myVariableNo = value;
            }
        }
        #endregion

        #endregion
    }
}
