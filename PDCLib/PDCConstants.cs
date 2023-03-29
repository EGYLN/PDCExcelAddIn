/// <summary>
/// ATTENTION: DO NOT EDIT, this is a generated resource.
/// Please edit Java class com.bayer.ws.pdc.service.PDCConstants instead.
/// </summary>
namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// This interface defines all PDC constants
  /// </summary>
  public class PDCConstants
  {
    /// <summary>
    /// Constant used for type conversion between date and string
    /// </summary>
    public const string C_FORMAT_DATE = "YYYY-MM-DD";

    /// <summary>
    /// Constant used for type conversion between timestamp and string
    /// </summary>
    public const string C_FORMAT_TIMESTAMP = "YYYY-MM-DD-HH24-MI-SS";

    /// <summary>
    /// Constant for PDC Parameter 'compoundtype'
    /// for distinction between compound no and preparaion no
    /// </summary>
    public const int C_ID_COMPOUNDTYPE = 1;

    /// <summary>
    /// Constant for PDC Parameter 'compoundidentifier'
    /// (either a compound no or preparation no.)
    /// </summary>
    public const int C_ID_COMPOUNDIDENTIFIER = 2;

    /// <summary>
    /// Constant for PDC Parameter 'test parameter'
    /// </summary>
    public const int C_ID_TEST_PARAMETER = 3;

    /// <summary>
    /// Constant for PDC Parameter 'test name'.
    /// </summary>
    public const int C_ID_TEST_NAME = 4;

    /// <summary>
    /// Constant for PDC Parameter 'version' used to identify the version of a test.
    /// </summary>
    public const int C_ID_VERSION = 5;

    /// <summary>
    /// Constant for PDC Parameter 'valid'. Currently unused
    /// </summary>
    public const int C_ID_VALID = 6;

    /// <summary>
    /// Constant for PDC Parameter 'molecular weight'
    /// </summary>
    public const int C_ID_WEIGHT = 7;

    /// <summary>
    /// Constant for PDC Parameter 'structure information'
    /// </summary>
    public const int C_ID_STRUCTURE_INFORMATION = 8;

    /// <summary>
    /// Constant for PDC Parameter 'projectname'
    /// </summary>
    public const int C_ID_PROJECTNAME = 9;

    /// <summary>
    /// Constant for PDC Parameter 'origin'. Output parameter to describe
    /// the origin of a test data set.
    /// </summary>
    public const int C_ID_ORIGIN = 10;

    /// <summary>
    /// Constant for PDC Parameter 'assay reference'
    /// </summary>
    public const int C_ID_ASSAY_REFERENCE = 11;

    /// <summary>
    /// Constant for PDC Parameter 'result status'
    /// </summary>
    public const int C_ID_RESULT_STATUS = 12;

    /// <summary>
    /// Constant for PDC Parameter 'date result'
    /// </summary>
    public const int C_ID_DATE_RESULT = 13;

    /// <summary>
    /// Constant for PDC Parameter 'reference'
    /// </summary>
    public const int C_ID_REFERENCE = 14;

    /// <summary>
    /// Constant for PDC Parameter 'preparationno'
    /// </summary>
    public const int C_ID_PREPARATIONNO = 15;

    /// <summary>
    /// Constant for PDC Parameter 'experimentno'. Will be generated and returned as output
    /// for every uploaded test and is used to identify a test during update and delete.
    /// </summary>
    public const int C_ID_EXPERIMENTNO = 16;

    /// <summary>
    /// Constant for PDC Parameter 'measurementno'. Used to distinct/relate measurements from/with each other.
    /// </summary>
    public const int C_ID_MEASUREMENTNO = 17;

    /// <summary>
    /// Constant for PDC Parameter 'variableno'. Used to relate a log (error) message to a certain variable no.
    /// </summary>
    public const int C_ID_VARIABLENO = 18;

    /// <summary>
    /// Constant for PDC Parameter 'file contents'. 
    /// Binary test parameters require input for file contents (blob), file name and format.
    /// Therefore they dont use the pdc parameter 'test parameter' as a special case.
    /// </summary>
    public const int C_ID_FILE_CONTENTS = 19;

    /// <summary>
    /// Constant for PDC Parameter 'file format'. 
    /// Binary test parameters require input for file contents (blob), file name and format.
    /// Therefore they dont use the pdc parameter 'test parameter' as a special case.
    /// </summary>
    public const int C_ID_FILE_FORMAT = 20;

    /// <summary>
    /// Constant for PDC Parameter 'file name'. 
    /// Binary test parameters require input for file contents (blob), file name and format.
    /// Therefore they dont use the pdc parameter 'test parameter' as a special case.
    /// </summary>
    public const int C_ID_FILE_NAME = 21;

    /// <summary>
    /// Constant for PDC Parameter 'DMSO amount'
    /// </summary>
    public const int C_ID_DMSO_AMOUNT = 22;

    /// <summary>
    /// Constant for PDC Parameter 'seperator'. Must be used to seperate multiple
    /// tests from each other.
    /// </summary>
    public const int C_ID_SEPERATOR = 23;

    /// <summary>
    /// Constant for PDC Parameter 'identifier'. Identifies a log message in the log table.
    /// </summary>
    public const int C_ID_IDENTIFIER = 24;

    /// <summary>
    /// Constant for PDC Parameter 'error date'. Specifies when an error occured.
    /// </summary>
    public const int C_ID_ERRORDATE = 25;

    /// <summary>
    /// Constant for PDC Parameter 'error'. An error message
    /// </summary>
    public const int C_ID_ERROR = 26;

    /// <summary>
    /// Constant for PDC Parameter 'message id'. Specifies the id of a message.
    /// </summary>
    public const int C_ID_MESSAGEID = 27;

    /// <summary>
    /// Constant for PDC Parameter 'source system'. 
    /// </summary>
    public const int C_ID_SOURCESYSTEM = 28;

    /// <summary>
    /// Constant for PDC Parameter 'runno'. Internally used to create an order.
    /// </summary>
    public const int C_ID_RUNNO = 29;

    /// <summary>
    /// Constant for PDC Parameter 'test no'
    /// </summary>
    public const int C_ID_TESTNO = 30;

    /// <summary>
    /// Constant for PDC Parameter 'person id'. Usually a cwid or sgno or name
    /// </summary>
    public const int C_ID_PERSONID = 31;

    /// <summary>
    /// Constant for PDC Parameter 'person id type'. Type of person id.
    /// </summary>
    public const int C_ID_PERSONID_TYPE = 32;

    /// <summary>
    /// Constant for PDC Parameter 'operation'. Internally used to distinct 
    /// insert and delete entries.
    /// </summary>
    public const int C_ID_OPERATION = 33;

    /// <summary>
    /// Constant for PDC Parameter 'alternate preparation no'
    /// </summary>
    public const int C_ID_ALTERNATE_PREPNO = 34;

    /// <summary>
    /// Constant for PDC Parameter 'replicated at'. 
    /// Internally used to identify replicated rows.
    /// </summary>
    public const int C_ID_REPLICATED_AT = 35;

    /// <summary>
    /// Constant for PDC Parameter 'upload date'.
    /// Internally used to set the date of an upload.
    /// </summary>
    public const int C_ID_UPLOADDATE = 36;

    /// <summary>
    /// Constant for PDC Parameter 'scheduled date'. Currently unused.
    /// </summary>
    public const int C_ID_SCHEDULEDDATE = 37;

    /// <summary>
    /// Constant for PDC Parameter 'overview only'. Used for finding test data.
    /// If the parameter is set only test identifying data is returned instead
    /// of the complete test data.
    /// </summary>
    public const int C_ID_OVERVIEW_ONLY = 38;

    /// <summary>
    /// Constant for PDC Parameter 'partial_compare'. Used for finding test data.
    /// Used to indicate that the uploaded test data may contain more data than 
    /// the caller has specified.
    /// </summary>
    public const int C_ID_PARTIAL_COMPARE = 39;

    /// <summary>
    /// Constant for PDC configuration parameter 'immediate transfer'. Used for Upload/Update/Delete
    /// of test data. Used to indicate that the new data should be replicated to PIx immediately.
    /// </summary>
    public const int C_ID_IMMEDIATE_TRANSFER = 40;

    /// <summary>
    /// PDC Parameter Constant for a name of a Testvariable
    /// </summary>
    public const int C_ID_VARIABLENAME = 41;

    /// <summary>
    /// PDC Parameter Constant for a variable type
    /// </summary>
    public const int C_ID_VARIABLETYPE = 42;

    /// <summary>
    /// PDC Parameter Constant for a variable class
    /// </summary>
    public const int C_ID_VARIABLECLASS = 43;

    /// <summary>
    /// PDC Parameter Constant for a variable unit
    /// </summary>
    public const int C_ID_UNIT = 44;

    /// <summary>
    /// PDC Parameter Constant for the is core result flag of a test variable
    /// </summary>
    public const int C_ID_ISCORERESULT = 45;

    /// <summary>
    /// PDC Parameter Constant for a comment
    /// </summary>
    public const int C_ID_COMMENT = 46;

    /// <summary>
    /// PDC Parameter Constant for a department no
    /// </summary>
    public const int C_ID_DEPARTMENTNO = 47;

    /// <summary>
    /// PDC Parameter Constant for a deparment
    /// </summary>
    public const int C_ID_DEPARTMENT = 48;

    /// <summary>
    /// PDC Parameter Constant for a source id
    /// </summary>
    public const int C_ID_SOURCEID = 49;

    /// <summary>
    /// PDC Parameter Constant for the curated flag
    /// </summary>
    public const int C_ID_CURATED = 50;

    /// <summary>
    /// PDC Parameter Constant for a cwid
    /// </summary>
    public const int C_ID_CWID = 51;

    /// <summary>
    /// PDC Parameter Constant for the technical id of a test version
    /// </summary>
    public const int C_ID_TESTVERSION_ID = 52;

    /// <summary>
    /// PDC Parameter Constant for the technical id of a test definition
    /// </summary>
    public const int C_ID_TEST_ID = 53;

    /// <summary>
    /// PDC Paraemter Constatn for the description of a test definition
    /// </summary>
    public const int C_ID_TEST_DESCRIPTION = 54;

    /// <summary>
    /// PDC Parameter Constant for the datetime of the last change
    /// </summary>
    public const int C_ID_DATE_CHANGE = 55;

    /// <summary>
    /// PDC Parameter Constant for the technical id of a privilege
    /// </summary>
    public const int C_ID_PRIVILEGE_ID = 56;

    /// <summary>
    /// PDC Parameter Constant for the name of a privilege
    /// </summary>
    public const int C_ID_PRIVILEGE_NAME = 57;

    /// <summary>
    /// PDC Parameter Constant for a log level
    /// </summary>
    public const int C_ID_LOG_LEVEL = 59;

    /// <summary>
    /// PDC Parameter Constant for the mc number
    /// </summary>
    public const int C_ID_MCNO = 60;

    /// <summary>
    /// PDC Parameter Constant for the test definition author
    /// </summary>
    public const int C_ID_TD_AUTHOR = 61;

    /// <summary>
    /// PDC Parameter Constant for the structure drawing
    /// </summary>
    public const int C_ID_STRUCTURE_DRAWING = 62;

    /// <summary>
    /// PDC Parameter Constant for the structure formula
    /// </summary>
    public const int C_ID_FORMULA = 63;

    /// <summary>
    /// PDC Parameter Constant for the maximum size of binary upload data
    /// </summary>
    public const int C_ID_BINARY_DATA_MAXSIZE = 64;

    /// <summary>
    /// PDC Parameter Constant for valid binary data formats
    /// </summary>
    public const int C_ID_BINARY_DATA_VALID_FORMAT = 65;

    /// <summary>
    /// PDC Parameter Constant for the technical id of color config entry
    /// </summary>
    public const int C_ID_COLOR_ID = 66;

    /// <summary>
    /// PDC Parameter Constant for the color config item name
    /// </summary>
    public const int C_ID_COLOR_ITEM = 67;

    /// <summary>
    /// PDC Parameter Constant for the red color component
    /// </summary>
    public const int C_ID_COLOR_RGB_R = 68;

    /// <summary>
    /// PDC Parameter Constant for the green color component
    /// </summary>
    public const int C_ID_COLOR_RGB_G = 69;

    /// <summary>
    /// PDC Parameter Constant for the blue color component
    /// </summary>
    public const int C_ID_COLOR_RGB_B = 70;

    /// <summary>
    /// PDC Parameter Constant for prefixes
    /// </summary>
    public const int C_ID_PREFIX = 71;

    /// <summary>
    /// PDC Parameter Constant for default values of test variables
    /// </summary>
    public const int C_ID_DEFAULT_VALUE = 72;

    /// <summary>
    /// PDC Parameter Constant for references to picklists
    /// </summary>
    public const int C_ID_PICKLIST_IDENTIFIER = 73;

    /// <summary>
    /// PDC Parameter Constant for pick list names
    /// </summary>
    public const int C_ID_PICKLIST_NAME = 74;

    /// <summary>
    /// PDC Parameter specifying the variable type of a pick list
    /// </summary>
    public const int C_ID_PICKLIST_TYPE = 75;

    /// <summary>
    /// PDC Parameter for the low limit of numeric pick lists
    /// </summary>
    public const int C_ID_PICKLIST_LOW_LIMIT = 76;

    /// <summary>
    /// PDC Parameter for the high limit of numeric pick lists
    /// </summary>
    public const int C_ID_PICKLIST_HIGH_LIMIT = 77;

    /// <summary>
    /// PDC Parameter for pick list values
    /// </summary>
    public const int c_id_picklist_value = 78;

    /// <summary>
    /// PDC Parameter for the technical id of a pdc user
    /// </summary>
    public const int C_ID_PDC_USER_ID = 79;

    /// <summary>
    /// PDC Parameter for the upload id
    /// </summary>
    public const int C_ID_UPLOAD_ID = 80;

    /// <summary>
    /// PDC Parameter for the technical id of a test variable
    /// </summary>
    public const int C_ID_VARIABLE_ID = 81;

    /// <summary>
    /// PDC Parameter for specifiying variables by their identifier instead of their variableno
    /// </summary>
    public const int C_ID_TESTPARAMETER_BY_ID = 82;

    /// <summary>
    /// Constant for PDC Parameter 'file contents' using the variableid
    /// </summary>
    public const int C_ID_FILE_CONTENTS_BY_ID = 83;

    /// <summary>
    /// Constant for PDC Parameter 'file format' using the variableid. 
    /// </summary>
    public const int C_ID_FILE_FORMAT_BY_ID = 84;

    /// <summary>
    /// Constant for PDC Parameter 'file name' using the variableid 
    /// </summary>
    public const int C_ID_FILE_NAME_BY_ID = 85;

    /// <summary>
    /// PDC Parameter for specifiying the table name of a predefined parameter
    /// </summary>
    public const int C_ID_PREDEF_TABLE = 86;

    /// <summary>
    /// PDC Parameter for specifiying the service name of a predefined parameter
    /// </summary>
    public const int C_ID_PREDEF_SERVICE = 87;

    /// <summary>
    /// PDC Parameter for the description of a predefined parameter
    /// </summary>
    public const int C_ID_PREDEF_DESCRIPTION = 88;

    /// <summary>
    /// PDC Parameter for specifiying a value of a reference table
    /// </summary>
    public const int C_ID_REFTABLE_VALUE = 89;

    /// <summary>
    /// PDC Parameter for the pixexperimentno
    /// </summary>
    public const int C_ID_PIXEXPERIMENTNO = 90;

    /// <summary>
    /// PDC Parameter for the upper limit of the upload date in a search. Should sometime be replaced by a more generic range search
    /// </summary>
    public const int C_ID_UPLOAD_DATE_UPPER_LIMIT = 91;

    /// <summary>
    /// PDC Parameter for the source id of imported upload data
    /// </summary>
    public const int C_ID_UPLOAD_SOURCE_ID = 92;

    /// <summary>
    /// PDC Parameter Differentiating specifies if a test variable is differentiating
    /// </summary>
    public const int C_ID_DIFFERENTIATING = 93;

    /// <summary>
    /// PDC Parameter Mandatary specifies if a test variable is mandatory
    /// </summary>
    public const int C_ID_MANDATORY = 94;

    /// <summary>
    /// PDC Parameter EXPERIMENTLEVEL specifies if a test variable is an experiment level parameter
    /// </summary>
    public const int C_ID_EXPERIMENTLEVEL = 95;

    /// <summary>
    /// PDC Version ID
    /// </summary>
    public const int C_ID_PDCVERSIONID = 96;

    /// <summary>
    /// Keep measurements
    /// </summary>
    public const int C_ID_KEEPMEASUREMENTS = 97;

    /// <summary>
    /// PDC Parameter EXPERIMENTLEVEL specifies if a test variable is an experiment level parameter
    /// </summary>
    public const int C_ID_COMPOUND_NO_EXP_LEVEL = 98;

    /// <summary>
    /// PDC Parameter EXPERIMENTLEVEL specifies if a test variable is an experiment level parameter
    /// </summary>
    public const int C_ID_PREPARATION_NO_EXP_LEVEL = 99;

    /// <summary>
    /// PDC Parameter EXPERIMENTLEVEL specifies if a test variable is an experiment level parameter
    /// </summary>
    public const int C_ID_MCNO_NO_EXP_LEVEL = 102;

    /// <summary>
    /// PDC Parameter PDCONLY specifies if the PDC 
    /// </summary>
    public const int C_ID_PDC_ONLY_DATA = 100;

    /// <summary>
    /// The Hydrogen parameter for the structure drawing generation
    /// </summary>
    public const int C_ID_HYDROGEN = 101;

    /// <summary>
    /// The parameter for deleted experiment nos during an update
    /// </summary>
    public const int C_ID_DELETED_EXPERIMENTNO = 103;

    /// <summary>
    /// Constant for compound type 'compoundno' enum value
    /// </summary>
    public const string C_TYPE_COMPOUND = "compound";

    /// <summary>
    /// Constant for compound type 'preparationno' enum value
    /// </summary>
    public const string C_TYPE_PREPARATION = "preparation";

    /// <summary>
    /// Constant for internally used parameter type 'string'
    /// </summary>
    public const int C_TYPE_STRING = 1;

    /// <summary>
    /// Constant for internally used parameter type 'number'
    /// </summary>
    public const int C_TYPE_NUMBER = 2;

    /// <summary>
    /// Constant for internally used parameter type 'integer'. Currently unused.
    /// </summary>
    public const int C_TYPE_INTEGER = 3;

    /// <summary>
    /// Constant for internally used parameter type 'double'. Currently unused.
    /// </summary>
    public const int C_TYPE_DOUBLE = 4;

    /// <summary>
    /// Constant for internally used parameter type 'string or number'. Currently unused.
    /// </summary>
    public const int C_TYPE_STRING_OR_NUMBER = 5;

    /// <summary>
    /// Constant for internally used parameter type 'blob'
    /// </summary>
    public const int C_TYPE_BLOB = 6;

    /// <summary>
    /// Constant for internally used parameter type 'boolean'
    /// </summary>
    public const int C_TYPE_BOOLEAN = 7;

    /// <summary>
    /// Constant for internally used parameter type 'date'
    /// </summary>
    public const int C_TYPE_DATE = 8;

    /// <summary>
    /// Constant for internally used parameter type 'timestamp'
    /// </summary>
    public const int C_TYPE_TIMESTAMP = 9;

    /// <summary>
    /// Constant for internally used parameter type 'test parameter'
    /// </summary>
    public const int C_TYPE_TESTPARAMETER = 10;

    /// <summary>
    /// Identification constant for message 'Invalid Parameter type'
    /// </summary>
    public const string C_MSG_INVALID_PARAM_TYPE = "PDC_PARTYPE";

    /// <summary>
    /// Identification constant for message 'Missing parameter'
    /// </summary>
    public const string C_MSG_MISSING_PARAM = "PDC_PARMISSING";

    /// <summary>
    /// Identification constant for message 'Invalid data type format'
    /// </summary>
    public const string C_MSG_INVALID_FORMAT = "PDC_FORMAT";

    /// <summary>
    /// Identification constant for message 'Unknown parameter'
    /// </summary>
    public const string C_MSG_UNKNOWN_PARAM = "PDC_PARUNKNOWN";

    /// <summary>
    /// Identification constant for message 'General SQL error'
    /// </summary>
    public const string C_MSG_SQL_ERROR = "PDC_SQLERROR";

    /// <summary>
    /// Identification constant for message 'Data missing'. Used when removing errors.
    /// if the number of found records does not match the number of provided identifiers.
    /// </summary>
    public const string C_MSG_PARTIAL_DELETE = "PDC_DATAMISSING";

    /// <summary>
    /// Identification constant for message 'Data not found'
    /// </summary>
    public const string C_MSG_DATA_NOT_FOUND = "PDC_DATANOTFOUND";

    /// <summary>
    /// Identification constant for message 'Value out of range' A provided value does not lie
    /// in the value range provided by the parameter constraints.
    /// </summary>
    public const string C_MSG_VALUE_OUT_OF_RANGE = "PDC_OUTOFRANGE";

    /// <summary>
    /// Identification constant for message 'Value not in enum'. A provided value does not match
    /// one of the enum values provided by the parameter constraints.
    /// </summary>
    public const string C_MSG_VALUE_NOT_IN_ENUM = "PDC_NOTINENUM";

    /// <summary>
    /// Identification constant for message 'Value missing'
    /// </summary>
    public const string C_MSG_VALUE_MISSING = "PDC_VALUEMISSING";

    /// <summary>
    /// Identification constant for message 'Unknown test parameter'
    /// </summary>
    public const string C_MSG_UNKNOWN_TEST_PARAM = "PDC_UNKNOWNTESTPARAM";

    /// <summary>
    /// Identification constant for message 'Invalid position'. Binary parameters are
    /// only possible as experiment level parameters.
    /// </summary>
    public const string C_MSG_INVALID_POSITION = "PDC_INVALIDPOSITION";

    /// <summary>
    /// Identification constant for message 'Value too large'
    /// </summary>
    public const string C_MSG_VALUE_TOO_LARGE = "PDC_VALUETOOLARGE";

    /// <summary>
    /// Identification constant for message 'Connection error'. Currently unused
    /// </summary>
    public const string C_MSG_CONNECTION_FAILED = "PDC_CONN_E";

    /// <summary>
    /// Identification constant for message 'Connection error'. Currently unused
    /// </summary>
    public const string C_MSG_PIX_CONNECTION_FAILED = "PDC_CONN_PIX_E";

    /// <summary>
    /// Identification constant for message 'Connection error'. Currently unused
    /// </summary>
    public const string C_MSG_ROOTS_CONNECTION_FAILED = "PDC_CONN_ROOTS_E";

    /// <summary>
    /// Identification constant for message 'Authentication error'. Currently unused.
    /// </summary>
    public const string C_MSG_AUTORIZATION_FAILED = "PDC_AUTH_E";

    /// <summary>
    /// Identification constant for message 'Internal error'
    /// </summary>
    public const string C_MSG_INTERNAL_ERROR = "PDC_INTERNAL";

    /// <summary>
    /// Identification constant for message 'Invalid Search Param'
    /// </summary>
    public const string C_MSG_INVALID_SEARCH_PARAM = "PDC_SEARCHPARAM";

    /// <summary>
    /// Identification constant for message 'Pix error'. Replicated from pix.
    /// </summary>
    public const string C_MSG_ERROR_FROM_PIX = "PDC_PIX_ERROR";

    /// <summary>
    /// Identification constant for message 'Invalid Binary'
    /// </summary>
    public const string C_MSG_INCOMPLETE_FILE_SPEC = "PDC_INVALIDBINARY";

    /// <summary>
    /// Identification constant for message 'Empty input'
    /// </summary>
    public const string C_MSG_EMPTY_INPUT = "PDC_EMPTY_INPUT";

    /// <summary>
    /// Identification constant for message 'Experiment param'. Test parameter is experiment level but
    /// has a position set.
    /// </summary>
    public const string C_MSG_EXPERIMENT_PARAM = "PDC_EXPERIMENT_PARAM";

    /// <summary>
    /// Identification constant for message 'Measurement param'. Test parameter is measurement level but
    /// no positive position is set.
    /// </summary>
    public const string C_MSG_MEASUREMENT_PARAM = "PDC_MEASUREMENT_PARAM";

    /// <summary>
    /// Identification constant for message 'Measurement param'. Test parameter is measurement level but
    /// no positive position is set.
    /// </summary>
    public const string C_MSG_REPLICATED = "PDC_REPLICATED";

    /// <summary>
    /// Identification constant for message 'No preparation no specified'
    /// </summary>
    public const string C_MSG_MISSING_PREPNO = "PDC_MISSING_PREPNO";

    /// <summary>
    /// Identification constant for message 'Multiple preparation no found'
    /// </summary>
    public const string C_MSG_MULTIPLE_PREPNO = "PDC_MULTIPLE_PREPNO";
    /// <summary>
    /// Identification constant for message 'Duplicate Parameter'. A parameter was defined twice with
    /// possibly different parameter values.
    /// </summary>
    public const string C_MSG_DUPLICATE_PARAM = "PDC_DUPLICATE_PARAM";

    /// <summary>
    /// Identification constant for warn message 'Unknown Compound no'. Test data for unknown compound no are
    /// not replicated to PIx.
    /// </summary>
    public const string C_MSG_UNKNOWN_COMPOUNDNO = "PDC_UNKNOWN_COMPOUND";

    /// <summary>
    /// Identification constant for warn message 'Unknown Preparationno'. Test data for unknown preparationnos are
    /// not replicated to PIx.
    /// </summary>
    public const string C_MSG_UNKNOWN_PREPNO = "PDC_UNKNOWN_PREPNO";

    /// <summary>
    /// Identification constant for warn message 'PDC_UNKNOWN_MCNO'.
    /// </summary>
    public const string C_MSG_UNKNOWN_MCNO = "PDC_UNKNOWN_MCNO";

    /// <summary>
    /// Identification constant for warn message 'PDC_MISSING_CNO'.
    /// </summary>
    public const string C_MSG_MISSING_CNO = "PDC_MISSING_CNO";

    /// <summary>
    /// Identification constant for warn message 'PDC_MISMATCH_CP'.
    /// </summary>
    public const string C_MSG_MISMATCH_CP = "PDC_MISMATCH_CP";

    public const string C_MSG_PDC_INVALIDC_VALIDP = "PDC_INVALIDC_VALIDP";

    public const string C_MSG_PDC_INVALIDP_VALIDM = "PDC_INVALIDP_VALIDM";

    public const string C_MSG_PDC_PREPARATION_COP = "PDC_PREPARATION_COP";

    /// <summary>
    /// Identification constant for warn message 'PDC_MISMATCH_PM'.
    /// </summary>
    public const string C_MSG_MISMATCH_PM = "PDC_MISMATCH_PM";

    /// <summary>
    /// Identification constant for into message 'PDC_COMPOUND_DEFERRED'.
    /// </summary>
    public const string C_MSG_PDC_COMPOUND_DEFERRED = "PDC_COMPOUND_DEFERRED";

    /// <summary>
    /// Identification constant for into message 'PDC_PREPARATION_DEFERRED'.
    /// </summary>
    public const string C_MSG_PDC_PREPARATION_DEFERRED = "PDC_PREPARATION_DEFERRED";

    /// <summary>
    /// Identification constant for warn message 'PDC_BAYMISSING'.
    /// </summary>
    public const string C_MSG_BAYMISSING = "PDC_BAYMISSING";

    /// <summary>
    /// Identification constant for warn message 'PDC_RESULTCOUNT'.
    /// </summary>
    public const string C_MSG_RESULT_COUNTER = "PDC_RESULTCOUNT";

    /// <summary>
    /// Identification constant for warn message 'PDC_SEARCHPARAMLIMIT'.
    /// </summary>
    public const string C_MSG_SEARCHPARAM_LIMIT = "PDC_SEARCHPARAMLIMIT";

    /// <summary>
    /// PDC Version mismatch
    /// </summary>
    public const string C_MSG_PDC_VERSION = "PDC_VERSION";

    /// <summary>
    /// Constant for log level ERROR
    /// </summary>
    public const int C_LOG_LEVEL_ERROR = 4;

    /// <summary>
    /// Constant for log level DEBUG
    /// </summary>
    public const int C_LOG_LEVEL_DEBUG = 1;

    /// <summary>
    /// Constant for log level FATAL
    /// </summary>
    public const int C_LOG_LEVEL_FATAL = 5;

    /// <summary>
    /// Constant for log level FINE
    /// </summary>
    public const int C_LOG_LEVEL_FINE = 2;

    /// <summary>
    /// Constant for log level INFO
    /// </summary>
    public const int C_LOG_LEVEL_INFO = 1;

    /// <summary>
    /// Constant for log level WARNING
    /// </summary>
    public const int C_LOG_LEVEL_WARNING = 3;

    /// <summary>
    /// Constant for Service method check compound name
    /// </summary>
    public const string C_OP_CHECK_COMPOUND_NAME = "check_compound_name";

    /// <summary>
    /// Constant for Service method list preparation numbers
    /// </summary>
    public const string C_OP_LIST_PREPARATION_NUMBERS = "list_preparation_numbers";

    /// <summary>
    /// Constant for Service method list projects
    /// </summary>
    public const string C_OP_LIST_PROJECTS = "list_projects";

    /// <summary>
    /// Constant for Service method return weight and structure
    /// </summary>
    public const string C_OP_WEIGHT_AND_STRUCTURE = "return_weight_and_structure";

    /// <summary>
    /// Constant for Service method calculate DMSO amount
    /// </summary>
    public const string C_OP_CALCULATE_DMSO_AMOUNT = "calculate_DMSO_amount";

    /// <summary>
    /// Constant for Service method check upload data exists
    /// </summary>
    public const string C_OP_CHECK_UPLOAD_DATA_EXISTS = "check_upload_data_exists";

    /// <summary>
    /// Constant for Service method validate data
    /// </summary>
    public const string C_OP_VALIDATE_DATA = "validate_data";

    /// <summary>
    /// Constant for Service method find upload data
    /// </summary>
    public const string C_OP_FIND_UPLOAD_DATA = "find_upload_data";

    /// <summary>
    /// Constant for Service method create error or message
    /// </summary>
    public const string C_OP_CREATE_ERROR_OR_MESSAGE = "create_error_or_message";

    /// <summary>
    /// Constant for Service method update error
    /// </summary>
    public const string C_OP_UPDATE_ERRORS = "update_error";

    /// <summary>
    /// Constant for Service method retrieve errors
    /// </summary>
    public const string C_OP_RETRIEVE_ERRORS = "retrieve_errors";

    /// <summary>
    /// Constant for Service method remove errors
    /// </summary>
    public const string C_OP_REMOVE_ERRORS = "remove_errors";

    /// <summary>
    /// Constant for Service method upload single entry
    /// </summary>
    public const string C_OP_UPLOAD_SINGLE_ENTRY = "upload_single_entry";

    /// <summary>
    /// Constant for Service method upload table
    /// </summary>
    public const string C_OP_UPLOAD_TABLE = "upload_table";

    /// <summary>
    /// Constant for Service method update table
    /// </summary>
    public const string C_OP_UPDATE_TABLE = "update_table";

    /// <summary>
    /// Constant for Service method delete table
    /// </summary>
    public const string C_OP_DELETE_TABLE = "delete_table";

    /// <summary>
    /// 
    /// </summary>
    public const string C_OP_FIND_TESTDEFINITION = "find_testdefinition";

    /// <summary>
    /// Constant for Service method get variables
    /// </summary>
    public const string C_OP_GET_VARIABLES = "get_variables";

    /// <summary>
    /// Constant for Service method get departments
    /// </summary>
    public const string C_OP_GET_DEPARTMENTS = "get_departments";

    /// <summary>
    /// Constant for Service method get_color_config
    /// </summary>
    public const string C_OP_GET_COLOR_CONFIGURATION = "get_color_config";

    /// <summary>
    /// Constant for Service method get_config
    /// </summary>
    public const string C_OP_GET_CONFIG = "get_config";

    /// <summary>
    /// Constant for Service method get_prefixes
    /// </summary>
    public const string C_OP_GET_PREFIXES = "get_prefixes";

    /// <summary>
    /// Constant for Service method check_pdcversion
    /// </summary>
    public const string C_OP_CHECK_VERSION = "check_pdcversion";

    /// <summary>
    /// Constant for Service method returning the picklists of a test version
    /// </summary>
    public const string C_OP_GET_PICKLISTS = "get_picklists";

    /// <summary>
    /// Constant for Service method which validates a collection of experiments
    /// </summary>
    public const string C_OP_VALIDATE_TABLE = "validate_table";

    /// <summary>
    /// Constant for Service method which returns the reference data from a specified table
    /// </summary>
    public const string C_OP_GET_REFERENCE_DATA = "get_reference_data";

    /// <summary>
    /// Constant for Service method which returns the definitions of all predefined parameters
    /// </summary>
    public const string C_OP_GET_PREDEFINED_PARAMETERS = "get_predefined_parameters";

    /// <summary>
    /// Constant for the Deletion of experiments by experimentnos
    /// </summary>
    public const string C_OP_DELETE_EXPERIMENTS = "delete_experiments";

    /// <summary>
    /// Constant for uploading workbook changes
    /// </summary>
    public const string C_OP_UPLOAD_CHANGES = "upload_changes";


    /// <summary>
    /// PDC Version
    /// </summary>
    public const string C_ID_PDCVERSION = "2.0";

    /// <summary>
    /// The technical id for test definition status 'PIX'
    /// </summary>
    public const int C_TD_STATUS_TYPE_PIX = 1;

    /// <summary>
    /// The technical id for test definition status 'outdated'
    /// </summary>
    public const int C_TD_STATUS_TYPE_OUTDATED = 5;

    /// <summary>
    /// The technical id for the Dynamic User Right 'Data Entry'
    /// </summary>
    public const int C_USERRIGHT_DATAENTRY = 2;

    /// <summary>
    /// The technical id for the result status type HTS
    /// </summary>
    public const int C_RESULTSTATUS_TYPE_HTS = 1;

    /// <summary>
    /// The technical id for the result status type Pharmacological
    /// </summary>
    public const int C_RESULTSTATUS_TYPE_PHARMA = 2;

    /// <summary>
    /// The technical id for the result status type active/inactive
    /// </summary>
    public const int C_RESULTSTATUS_TYPE_ACTIVE = 3;

    /// <summary>
    /// The technical id for the result status type differentiating
    /// </summary>
    public const int C_RESULTSTATUS_TYPE_DIFF = 4;

    /// <summary>
    /// Representation of a missing BAY compound no
    /// </summary>
    public const string C_BAY_MISSING = "BAY MISSING";

    /// <summary>
    /// Number of days BAY_MISSING is tried to be replaced
    /// </summary>
    public const int C_BAY_MISSING_TIMER = 15;

    /// <summary>
    /// Max allowed result count for findUploadData()
    /// </summary>
    public const int C_RESULT_COUNTER_LIMIT = 20000;

    /// <summary>
    /// Constant for Result status Value 'No Effect'
    /// </summary>
    public const int RESULT_STATUS_NO_EFFECT = 0;

    /// <summary>
    /// Constant for Result status Value 'Effect'
    /// </summary>
    public const int RESULT_STATUS_EFFECT = 1;

    /// <summary>
    /// Constant for Person Type CWID
    /// </summary>
    public const int PERSON_TYPE_ID = 1;
    

  }
}
