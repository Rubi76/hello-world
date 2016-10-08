using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Collections;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

#region NAMESPACE VMS.TPS
namespace VMS.TPS
{

    #region ----- SCRIPT TPS ------
    /// <summary>
    /// Scripting Plan Check
    /// </summary>
    public class Script
    {
        /// <summary>
        /// Script constructor without paramameters.
        /// </summary>
        public Script()
        {
        }

        /// <summary>
        /// Information Table messages definition
        /// </summary>
        public static DataTable informationTable = new DataTable();
        /// <summary>
        /// Verification Table messages definition
        /// </summary>
        public static DataTable verificationTable = new DataTable();
        /// <summary>
        /// Warning Table messages definition
        /// </summary>
        public static DataTable warningTable = new DataTable();
       

        /// <summary>
        /// Script (Main Program) 
        /// </summary>
        /// <param name="context">
        /// The TPS context.
        /// </param>
        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {
            // ----------------- Context variables ------------------------                        
            PlanSetup plan = context.PlanSetup;
            Patient patient = context.Patient;
            Image imge = context.Image;
            Course course = context.Course;
            LocalScriptContext localContext = null;


            // --------------  Initial checks  -----------------------------    

            string initialMsg = string.Empty;
            if (PlanCheckInitialVerification.CheckInitialParameters(context, plan, imge, out initialMsg))
            {
                throw new Exception(initialMsg); //Exit Script
            }

            // Local Context generation
            try
            {
                localContext = LocalScriptContextGeneration.localScriptContextGeneration();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            /// ----------------- Creating headers of message tables --------------            
            // Information messages table            
            Script.informationTable.Columns.Add("informationTable", typeof(string));

            // Verification messages table           
            Script.verificationTable.Columns.Add("verificationMessage", typeof(string));
            Script.verificationTable.Columns.Add("booleanBool", typeof(bool));

            // Warning messages table            
            Script.warningTable.Columns.Add("warningMessage", typeof(string));



            //// ------------------ Information ------------------
            try
            {
                CheckInformation.GetInformation(plan, imge);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            //------------------ Verification and warnings ------------------   
            try
            {
                CheckVerification.VerificationTests(plan, imge, localContext);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            // ---------------------------------------------------------------------
            // -----------------------  Printing -----------------------------------
            // ---------------------------------------------------------------------
            //"\u2714" is the "ok" symbol in unicode
            //"\u2718" is the "x" symbol in unicode
            //"\u0020" is the "space" symbol in unicode
            //"\u2022" is the "bullet" symbol in unicode

            // ---------------  Printing information, verification and warning tables ------------
            try
            {
                CheckPrinting.PrintInformationMessage(informationTable);
                CheckPrinting.PrintVerifMessage(verificationTable);
                CheckPrinting.PrintWarningMessage(warningTable);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            // --------- Close SQL connection ---------------
            try
            {
                CheckSQL.SQLCloseConnection(localContext.MySQLConnection.Conn);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

    }

    // ---------------------------------- End Script class ------------------------------------------------
    #endregion

    #region LOCAL CONTEXT

    #region LOCAL SCRIPT CONTEXT CLASSES
    /// <summary>
    /// Defining the local context.
    /// </summary>
    public class LocalScriptContext
    {
        /// <summary>
        /// Local context constructor.
        /// </summary>
        /// <param name="localContext"> Tuple with classes in local context.</param>
        public LocalScriptContext(Tuple<MachinesList, MySQLConnection, Calculation, Collision> localContext)
        {
            if (localContext.Item1 is MachinesList)     //List of Machines with their definitions.
            {
                Machines = localContext.Item1;
            }
            if (localContext.Item2 is MySQLConnection)  //SQL paramenters connection.
            {
                MySQLConnection = localContext.Item2;
            }
            if (localContext.Item3 is Calculation)      //Calculation parameters
            {
                Calculation = localContext.Item3;
            }
            if (localContext.Item4 is Collision)        //Collision gantry-table -- gantry-patient parameters
            {
                Collision = localContext.Item4;
            }

        }

        /// <summary>
        /// Define local Machines
        /// </summary>
        public MachinesList Machines { get; private set; } //only reading
        /// <summary>
        /// Define SQL connection
        /// </summary>
        public MySQLConnection MySQLConnection { get; private set; } //only reading
        /// <summary>
        /// Define local calculation parameters
        /// </summary>
        public Calculation Calculation { get; private set; } //only reading
        /// <summary>
        /// Define local gantry collision parameters
        /// </summary>
        public Collision Collision { get; private set; } //only reading

    }

    /// <summary>
    /// Local machines list
    /// </summary>
    public class MachinesList : IEnumerable<Machine>
    {
        private List<Machine> _elements;

        /// <summary>
        /// Local machine list
        /// </summary>
        /// <param name="array"> List of local machines</param>
        public MachinesList(Machine[] array)
        {
            this._elements = new List<Machine>(array);
        }

        /// <summary>
        /// Implementation for the GetEnumerator method to get the local machine object.
        /// </summary>
        /// <returns>Local machine object</returns>
        IEnumerator<Machine> IEnumerable<Machine>.GetEnumerator()
        {
            return this._elements.GetEnumerator();
        }

        /// <summary>
        /// Implementation of default GetEnumerator method.
        /// </summary>
        /// <returns>Element of interface Enumerable</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this._elements.GetEnumerator();
        }

    }

    /// <summary>
    /// Local machine
    /// </summary>
    public class Machine
    {
        private string _MachineId;
        private Couch _Couch;

        /// <summary>
        /// Definition of the local machine ID property.
        /// </summary>
        public string MachineId
        {
            get { return _MachineId; }
            set { _MachineId = value; }
        }

        /// <summary>
        /// Definition of the machine couch property
        /// </summary>
        public Couch Couch
        {
            get { return _Couch; }
            set { _Couch = value; }
        }

    }

    public class Couch
    {
        private string _CouchName;
        private List<CouchRegion> _CouchRegion;

        /// <summary>
        /// Definition of the local couch name property.
        /// </summary>
        public string CouchName
        {
            get { return _CouchName; }
            set { _CouchName = value; }
        }

        /// <summary>
        /// List of couch regions (parts)
        /// </summary>
        public List<CouchRegion> CouchRegions
        {
            get { return _CouchRegion; }
            set { _CouchRegion = value; }
        }
    }

    /// <summary>
    /// Couch regions definition (Example: Thin, medium, thick)
    /// </summary>
    public class CouchRegion
    {
        private string _CouchRegionName;                // Equivalent in ScriptContext -> Structure.Name
        private double _CouchVertPositionCorrection;
        private List<CouchPart> _couchParts;
        private List<Tuple<int, int, int>> _CouchCollisionParameters;

        /// <summary>
        /// Definition of the couch region name property
        /// </summary>
        public string CouchRegionName
        {
            get { return _CouchRegionName; }
            set { _CouchRegionName = value; }
        }

        /// <summary>
        /// Definition of the vertical couch position correction property
        /// </summary>
        public double CouchVertPositionCorrection
        {
            get { return _CouchVertPositionCorrection; }
            set { _CouchVertPositionCorrection = value; }
        }

        /// <summary>
        /// List with the couch collision parameters.
        /// </summary>
        public List<Tuple<int, int, int>> CouchCollisionParameters
        {
            get { return _CouchCollisionParameters; }
            set { _CouchCollisionParameters = value; }
        }

        /// <summary>
        /// List with the couch parts.
        /// </summary>
        public List<CouchPart> CouchParts
        {
            get { return _couchParts; }
            set { _couchParts = value; }
        }
    }

    /// <summary>
    /// Couch regions definition (Example: CouchSurfade, CouchInterior)
    /// </summary>
    public class CouchPart
    {
        private string _CouchPieceId;       // Equivalent in ScriptContext -> Structure.Id
        private double _HU;                 // Equivalent in ScriptContext -> Structure.GetAssignedHU (out huValue)

        /// <summary>
        /// Definition of the couch part ID correction property
        /// </summary>
        public string CouchPieceId
        {
            get { return _CouchPieceId; }
            set { _CouchPieceId = value; }
        }

        /// <summary>
        /// Definition of the Hounsfield Units property for the couch part.
        /// </summary>
        public double HU
        {
            get { return _HU; }
            set { _HU = value; }
        }
    }

    /// <summary>
    /// SQL connection
    /// </summary>
    public class MySQLConnection
    {
        private string _dataSource;
        private string _initialCatalog;
        private string _userID;
        private string _password;
        private int _timeOut;
        private SqlConnection _conn;

        /// <summary>
        /// DataSource property to connect to SQL Server.
        /// </summary>
        public string DataSource
        {
            get { return _dataSource; }
            set { _dataSource = value; }
        }

        /// <summary>
        /// InitialCatalog property to connect to SQL Server.
        /// </summary>
        public string InitialCatalog
        {
            get { return _initialCatalog; }
            set { _initialCatalog = value; }
        }

        /// <summary>
        /// UserID property to connect to SQL Server.
        /// </summary>
        public string UserID
        {
            get { return _userID; }
            set { _userID = value; }
        }

        /// <summary>
        /// Password property to connect to SQL Server.
        /// </summary>
        public string Password
        {
            get { return _password; }
            set { _password = value; }
        }

        /// <summary>
        /// Time(in seconds) property to wait while trying to establish a connection 
        /// before terminating the attempt and generating an error.
        /// </summary>
        public int TimeOut
        {
            get { return _timeOut; }
            set { _timeOut = value; }
        }

        /// <summary>
        /// SQL connection
        /// </summary>
        public SqlConnection Conn
        {
            get { return _conn; }
            set { _conn = value; }
        }
    }

    /// <summary>
    /// Local calculation properties
    /// </summary>
    public class Calculation
    {
        private string _photonVolumeDoseAlgorithm;
        private string _electronVolumeDoseAlgorithm;
        private double _minMeanTargetDose;
        private double _maxMeanTargetDose;
        private List<string> _technique;
        private List<Tuple<string, Int32>> _techniqueRR;
        private Int32 _MUForRR6;

        /// <summary>
        /// Photon Volume Dose Algorithm property
        /// </summary>
        public string PhotonVolumeDoseAlgorithm
        {
            get { return _photonVolumeDoseAlgorithm; }
            set { _photonVolumeDoseAlgorithm = value; }
        }

        /// <summary>
        /// Electron Volume Dose Algorithm property
        /// </summary>
        public string ElectronVolumeDoseAlgorithm
        {
            get { return _electronVolumeDoseAlgorithm; }
            set { _electronVolumeDoseAlgorithm = value; }
        }

        /// <summary>
        /// Minimum Mean Target Dose advisable property
        /// </summary>
        public double MinMeanTargetDose
        {
            get { return _minMeanTargetDose; }
            set { _minMeanTargetDose = value; }
        }

        /// <summary>
        /// Maximum Mean Target Dose advisable property
        /// </summary>
        public double MaxMeanTargetDose
        {
            get { return _maxMeanTargetDose; }
            set { _maxMeanTargetDose = value; }
        }

        /// <summary>
        /// List of treatment Technique (Ex: VMAT, IMRT, ARC, 3D)
        /// </summary>
        public List<string> Techniques
        {
            get { return _technique; }
            set { _technique = value; }
        }

        /// <summary>
        /// Minimum rep rate for different techniques.
        /// </summary>
        public List<Tuple<string, Int32>> TechniqueRR
        {
            get { return _techniqueRR; }
            set { _techniqueRR = value; }
        }

        /// <summary>
        /// Number of MU to advise using 600 UM/min
        /// </summary>
        public Int32 MUForRR6
        {
            get { return _MUForRR6; }
            set { _MUForRR6 = value; }
        }
    }

    /// <summary>
    /// Gantry collision calculation 
    /// </summary>
    public class Collision
    {
        private double _maxCouchRotCalc;
        private double _maxCouchRotWarning;
        private double _collisionSafetyMarginDistance;
        private double _collisionSafetyMarginGantryAngle;
        private double _couchVertPositionCTCorrection;
        /// <summary>
        /// Maximum couch rotation property for the gantry collision calculation.
        /// </summary>
        public double MaxCouchRotCalc
        {
            get { return _maxCouchRotCalc; }
            set { _maxCouchRotCalc = value; }
        }
        /// <summary>
        /// Maximum couch rotation property for warning message.
        /// </summary>
        public double MaxCouchRotWarning
        {
            get { return _maxCouchRotWarning; }
            set { _maxCouchRotWarning = value; }
        }
        /// <summary>
        /// Collision safety margin distance property.
        /// </summary>
        public double CollisionSafetyMarginDistance
        {
            get { return _collisionSafetyMarginDistance; }
            set { _collisionSafetyMarginDistance = value; }
        }
        /// <summary>
        /// Collision safety margin gantry angle property.
        /// </summary>
        public double CollisionSafetyMarginGantryAngle
        {
            get { return _collisionSafetyMarginGantryAngle; }
            set { _collisionSafetyMarginGantryAngle = value; }
        }
        /// <summary>
        /// Couch vertical position CT correction property.
        /// </summary>
        public double CouchVertPositionCTCorrection
        {
            get { return _couchVertPositionCTCorrection; }
            set { _couchVertPositionCTCorrection = value; }
        }
    }

    #endregion

    #region LOCAL SCRIPT CONTEXT GENERATION
    /// <summary>
    /// Local Script Context Generation.
    /// </summary>
    public class LocalScriptContextGeneration
    {
        /// <summary>
        /// Generation of the local context.
        /// </summary>
        /// <returns> Local Context </returns>
        public static LocalScriptContext localScriptContextGeneration()
        {

            /// ---------- COUCHES -----------

            // -- Exact IGRT Couch --
            Couch couchExactIGRT = new Couch();
            couchExactIGRT.CouchName = "IGRTCouch";
            List<CouchRegion> LstCouchRegion = new List<CouchRegion>();
            CouchRegion couchRegion = new CouchRegion();
                couchRegion.CouchRegionName = "Exact IGRT Couch, medium";
                couchRegion.CouchVertPositionCorrection = 31.0;
                List<Tuple<int, int, int>> couchCollParamList = new List<Tuple<int, int, int>>();         //{a, b, r}  -> a= couchHigh; b = half couch width; r = radio free gantry collision.
                Tuple<int, int, int> couchCollParamTuple = new Tuple<int, int, int>(70, 215, 395);        // son valores para thick
                couchCollParamList.Add(couchCollParamTuple);
                couchCollParamTuple = new Tuple<int, int, int>(20, 270, 395);                           // son valores para thick
                couchCollParamList.Add(couchCollParamTuple);
                couchRegion.CouchCollisionParameters = couchCollParamList;
                List<CouchPart> LstCouchPartCouchExactIGRT = new List<CouchPart>();
                CouchPart couchPart = new CouchPart();
                couchPart.CouchPieceId = "CouchInterior";
                couchPart.HU = -1000;
                LstCouchPartCouchExactIGRT.Add(couchPart);
                couchPart = new CouchPart();
                couchPart.CouchPieceId = "CouchSurface";
                couchPart.HU = -300;
                LstCouchPartCouchExactIGRT.Add(couchPart);
                couchRegion.CouchParts = LstCouchPartCouchExactIGRT;
                LstCouchRegion.Add(couchRegion);

            couchRegion = new CouchRegion();
                couchRegion.CouchRegionName = "Exact IGRT Couch, thick";
                couchRegion.CouchVertPositionCorrection = 36.1;
                couchCollParamList = new List<Tuple<int, int, int>>();
                couchCollParamTuple = new Tuple<int, int, int>(70, 215, 395);
                couchCollParamList.Add(couchCollParamTuple);
                couchCollParamTuple = new Tuple<int, int, int>(20, 270, 395);
                couchCollParamList.Add(couchCollParamTuple);
                couchRegion.CouchCollisionParameters = couchCollParamList;
                couchRegion.CouchParts = LstCouchPartCouchExactIGRT;
                LstCouchRegion.Add(couchRegion);

            couchRegion = new CouchRegion();
                couchRegion.CouchRegionName = "Exact IGRT Couch, thin";
                couchRegion.CouchVertPositionCorrection = 26.6;
                couchCollParamList = new List<Tuple<int, int, int>>();
                couchCollParamTuple = new Tuple<int, int, int>(70, 215, 395);  // son valores para thick
                couchCollParamList.Add(couchCollParamTuple);
                couchCollParamTuple = new Tuple<int, int, int>(20, 270, 395);  // son valores para thick
                couchCollParamList.Add(couchCollParamTuple);
                couchRegion.CouchCollisionParameters = couchCollParamList;
                couchRegion.CouchParts = LstCouchPartCouchExactIGRT;
                LstCouchRegion.Add(couchRegion);

            couchExactIGRT.CouchRegions = LstCouchRegion;

            // -- Exact Couch --
            Couch couchExactCouch = new Couch();
            couchExactCouch.CouchName = "ExactCouch";
            LstCouchRegion = new List<CouchRegion>();
            couchRegion = new CouchRegion();
                couchRegion.CouchRegionName = "Exact Couch with Flat panel";
                couchRegion.CouchVertPositionCorrection = 13.8;
                couchCollParamList = new List<Tuple<int, int, int>>();
                couchCollParamTuple = new Tuple<int, int, int>(110, 240, 400);        // son valores para CouchRail en OUT
                couchCollParamList.Add(couchCollParamTuple);
                couchRegion.CouchCollisionParameters = couchCollParamList;
                LstCouchPartCouchExactIGRT = new List<CouchPart>();
                couchPart = new CouchPart();
                couchPart.CouchPieceId = "CouchRailLeft";
                couchPart.HU = 200;
                LstCouchPartCouchExactIGRT.Add(couchPart);
                couchPart = new CouchPart();
                couchPart.CouchPieceId = "CouchRailRight";
                couchPart.HU = 200;
                LstCouchPartCouchExactIGRT.Add(couchPart);
                couchRegion.CouchParts = LstCouchPartCouchExactIGRT;
                LstCouchRegion.Add(couchRegion);

            couchRegion = new CouchRegion();
                couchRegion.CouchRegionName = "Exact Couch with Unipanel, large window";
                couchRegion.CouchVertPositionCorrection = 14.4;
                couchCollParamList = new List<Tuple<int, int, int>>();
                couchCollParamTuple = new Tuple<int, int, int>(110, 240, 400);
                couchCollParamList.Add(couchCollParamTuple);
                couchRegion.CouchCollisionParameters = couchCollParamList;
                couchRegion.CouchParts = LstCouchPartCouchExactIGRT;
                LstCouchRegion.Add(couchRegion);

            couchRegion = new CouchRegion();
                couchRegion.CouchRegionName = "Exact Couch with Unipanel, small windows";
                couchRegion.CouchVertPositionCorrection = 13.5;
                couchCollParamList = new List<Tuple<int, int, int>>();
                couchCollParamTuple = new Tuple<int, int, int>(110, 240, 400);  // son valores para thick
                couchCollParamList.Add(couchCollParamTuple);
                couchRegion.CouchCollisionParameters = couchCollParamList;
                couchRegion.CouchParts = LstCouchPartCouchExactIGRT;
                LstCouchRegion.Add(couchRegion);

            couchExactCouch.CouchRegions = LstCouchRegion;

            /// -------- MACHINES -------- 
            var listOfMachine = new List<Machine>();

            // -- Trilogy --
            Machine trilogy = new Machine();
            trilogy.MachineId = "Trilogy";
            trilogy.Couch = couchExactIGRT;
            listOfMachine.Add(trilogy);

            // -- iX --
            Machine iX = new Machine();
            iX.MachineId = "iX";
            iX.Couch = couchExactIGRT;
            listOfMachine.Add(iX);

            // -- URTE --
            Machine URTE = new Machine();
            URTE.MachineId = "URTE";
            URTE.Couch = couchExactCouch;
            listOfMachine.Add(URTE);

            // -- 2100CD - gegant --
            Machine CD2100 = new Machine();
            CD2100.MachineId = "2100CD";
            CD2100.Couch = couchExactCouch;
            listOfMachine.Add(CD2100);

            Machine[] arrayOfMachine = listOfMachine.ToArray();
            MachinesList machines = new MachinesList(arrayOfMachine);

            /// -------- SQLCONNECTION ---------
            MySQLConnection SQLconn = new MySQLConnection();
            SQLconn.DataSource = "VARIANSRV";
            SQLconn.InitialCatalog = "variansystem";
            SQLconn.UserID = "reports";
            SQLconn.Password = "reports";
            SQLconn.TimeOut = 10;
            SQLconn.Conn = CheckSQL.SQLOpenConnection(SQLconn.DataSource, SQLconn.InitialCatalog, SQLconn.UserID, SQLconn.Password, SQLconn.TimeOut);

            /// --------- CALCULATION ------------ 
            Calculation calculation = new Calculation();
            calculation.ElectronVolumeDoseAlgorithm = "GGPB_10028";
            calculation.PhotonVolumeDoseAlgorithm = "AAA_13535";
            calculation.MinMeanTargetDose = 0.98;
            calculation.MaxMeanTargetDose = 1.02;
            calculation.MUForRR6 = 300;
            List<string> techniquesList = new List<string>();
            string techniques = "VMAT";
            techniquesList.Add(techniques);
            techniques = "ARC";
            techniquesList.Add(techniques);
            techniques = "IMRT";
            techniquesList.Add(techniques);
            techniques = "3D";
            techniquesList.Add(techniques);
            calculation.Techniques = techniquesList;
            List<Tuple<string, Int32>> techRRList = new List<Tuple<string, Int32>>();
            Tuple<string, Int32> techRRTuple = new Tuple<string, Int32>(calculation.Techniques[0], 600); //VMAT
            techRRList.Add(techRRTuple);
            techRRTuple = new Tuple<string, Int32>(calculation.Techniques[2], 400); //IMRT
            techRRList.Add(techRRTuple);

            calculation.TechniqueRR = techRRList;

            /// --------- COLLISION ----------- 
            Collision collision = new Collision();
            collision.MaxCouchRotCalc = 10; // degrees
            collision.MaxCouchRotWarning = 10; //degrees
            collision.CouchVertPositionCTCorrection = 69.3; //mm
            collision.CollisionSafetyMarginDistance = 10; //mm
            collision.CollisionSafetyMarginGantryAngle = 2; // degrees        

            /// --------- LOCAL CONTEXT GENERATION -----------
            Tuple<MachinesList, MySQLConnection, Calculation, Collision> tuplaObject = new Tuple<MachinesList, MySQLConnection, Calculation, Collision>(machines, SQLconn, calculation, collision);
            LocalScriptContext localContext = new LocalScriptContext(tuplaObject);
            return localContext;
        }
    }
    #endregion

    #endregion

    #region GENERAL INFORMATION
    /// <summary>
    /// General Information
    /// </summary>
    public class CheckInformation
    {
        /// <summary>
        /// Get general information and add into information table.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="imge">
        /// Represents a CT dataset. 
        /// </param>
        public static void GetInformation(PlanSetup plan, Image imge)
        {
            string infMeanTargetDose = null;
            try
            {
                infMeanTargetDose = CheckDoseVolume.GetMeanTargetDose(plan, "Relative").ToString();
            }
            catch (Exception)
            {
                infMeanTargetDose = ("N/A");
            }

            try
            {               
                Script.informationTable.Rows.Add("Plan :\t" + plan.Id);
                Script.informationTable.Rows.Add();
                Script.informationTable.Rows.Add("Total Dose:\t" + CheckPlanPrescription.GetTotalDose(plan).ToString());
                Script.informationTable.Rows.Add("Dose/fraction :\t" + CheckPlanPrescription.GetDosePerFracction(plan).ToString());
                Script.informationTable.Rows.Add();
                Script.informationTable.Rows.Add("Target volume:\t" + plan.TargetVolumeID);
                Script.informationTable.Rows.Add("     Mean Dose:\t" + infMeanTargetDose);
                Script.informationTable.Rows.Add("     Coverage V95 :\t" + CheckDoseVolume.GetPercentTargetCoverage(plan, 95).ToString() + " %");
                Script.informationTable.Rows.Add();
                Script.informationTable.Rows.Add("Body:");
                Script.informationTable.Rows.Add("     Dmax :\t" + CheckDoseVolume.GetMaxBodyDose(plan));
                Script.informationTable.Rows.Add("     Dmax(2cm3) :\t" + CheckDoseVolume.GetRelativeDoseAtVolume(2.0, plan, CheckStructure.GetBodyStructure(plan)));
                Script.informationTable.Rows.Add();
                Script.informationTable.Rows.Add("Number of Isos :\t" + CheckPlan.GetNIsos(plan, imge));
                Script.informationTable.Rows.Add();
                Script.informationTable.Rows.Add("MU total:\t" + CheckPlanCalculation.GetMUTotal(plan).ToString() + " MU");
                Script.informationTable.Rows.Add("MU/Gy :\t" + Math.Round(CheckPlanCalculation.GetPlanMUperGy(plan), 0).ToString() + " MU/Gy");
               
            }
            catch (Exception e)
            {
                throw new Exception("GetInformation function Error: \n" + e.Message);
            }

        }

    }
    #endregion

    #region GENERAL TESTS
    /// <summary>
    /// General verification test
    /// </summary>
    public class CheckVerification {
        /// <summary>
        /// Run all verification tests.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="imge">
        /// Represents a CT dataset. 
        /// </param>
        /// <param name="localContext">
        /// Represents local context.
        /// </param>
        public static void VerificationTests(PlanSetup plan, Image imge, LocalScriptContext localContext)
        {
            try
            {
                string msg = string.Empty;
                CheckPlanPrescription.CheckDoseRTPrescription(plan, localContext);  //Check if Planned Dose Prescription matches RT Prescription (Phycisian's Intent). 
                CheckDoseVolume.CheckTargetMeanDose(plan, localContext);
                CheckImage.CheckUserOriginModified(imge);
                CheckPlan.CheckSetupField(plan);
                CheckPlan.CheckAllFieldsSameMachine(plan);
                CheckDoseVolume.CheckBodyDoseCoverage(plan);
                CheckBeam.Check600RepRateAdvice(plan, localContext);
            }
            catch (Exception e)
            {
                throw new Exception("VerificationTests function Error: \n" + e.Message);
            }

        }

    }
    #endregion

    #region PLAN FUNCTIONS

    #region PLAN CHECK INITIAL VERIFICATION FUNCTION

    /// <summary>
    /// Set of functions to verify plan.
    /// </summary>
    public class PlanCheckInitialVerification
    {
        /// <summary>
        /// Control of the basic parameters in order that the script works.
        /// </summary>
        /// <param name="context">
        /// The TPS context.
        /// </param>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="imge">
        /// Represents a CT dataset. 
        /// </param>
        /// <param name="msg">
        /// Out string parameter with the wrong parameters.
        /// </param>
        /// <returns>
        /// True if some basic parameter is not correct.
        /// </returns>
        public static bool CheckInitialParameters(ScriptContext context, PlanSetup plan, Image imge, out string msg)
        {

            bool error = false;
            msg = string.Empty;
            try
            {
                // If there's no selected plan with calculated dose throw an exception
                if (plan == null && context.PlanSumsInScope.Count() > 0)
                {
                    msg += ("\u2022   Script not valid for plan sum.\n");
                    error = true;
                }

                if (plan != null)
                {
                    if (plan.Dose == null)
                    {
                        msg += ("\u2022   Please open a calculated plan before using this script.\n");
                        error = true;
                    }


                    if (plan.TargetVolumeID == null | plan.TargetVolumeID == string.Empty)
                    {
                        msg += ("\u2022   Please a target volume is necessary before using this script.\n");
                        error = true;
                    }

                    try
                    {
                        CheckPlanPrescription.GetNumberOfFracctions(plan);
                        CheckPlanPrescription.GetDosePerFracction(plan);
                        CheckPlanPrescription.GetTotalDose(plan);
                    }
                    catch (Exception)
                    {
                        msg += ("\u2022   Check the Plan Dose prescription.\n");
                        error = true;
                    }
                }

                if (imge == null)
                {
                    msg += ("\u2022   There is no CT image linked to the plan.\n");
                    error = true;
                }

                return error;

            }
            catch (Exception e)
            {
                throw new Exception("CheckInitialParameters function Error: \n" + e.Message);
            }
        }


    }

    //  ---------------------------   End Verification Functions --------------------------------------

    #endregion

    #region PLAN GENERAL FUNCTIONS

    /// <summary>
    /// General Plan Functions.
    /// </summary>
    public class CheckPlan
    {
        /// <summary>
        /// Get the number of isos.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="imge">
        /// Represents a CT dataset. 
        /// </param>
        /// <returns>
        /// Number of isos
        /// </returns>
        public static int GetNIsos(PlanSetup plan, Image imge)
        {
            try
            {
                VVector[] vVectorArrayIsoDicomPos = new VVector[CheckPlan.GetNumberOfBeams(plan)];
                VVector[] vVectorArrayIsos = new VVector[CheckPlan.GetNumberOfBeams(plan)];
                int i = 0;

                //string message = "";

                // Get the Dicom position of isocenter for each field.
                foreach (Beam b in plan.Beams)
                {
                    //message += ("b.IsocenterPosition.x: " + b.IsocenterPosition.x + "\n");
                    //message += ("imge.UserOrigin.x: " + imge.UserOrigin.x);

                    if (!Double.IsNaN(b.IsocenterPosition.x))
                    {
                        vVectorArrayIsoDicomPos[i] = b.IsocenterPosition;
                        i++;
                        //MessageBox.Show(b.IsocenterPosition.x.ToString() + " , " + b.IsocenterPosition.y.ToString() + " , " + b.IsocenterPosition.z.ToString());
                    }
                }
                // MessageBox.Show(message);

                // Get the position of isocenter for each field from User Origin Dicom offset, that is to say displacements.
                for (int n = 0; n < vVectorArrayIsoDicomPos.Length; n++)
                {
                    vVectorArrayIsos[n].x = Math.Round(vVectorArrayIsoDicomPos[n].x - imge.UserOrigin.x, 2);
                    vVectorArrayIsos[n].y = Math.Round(vVectorArrayIsoDicomPos[n].y - imge.UserOrigin.y, 2);
                    vVectorArrayIsos[n].z = Math.Round(vVectorArrayIsoDicomPos[n].z - imge.UserOrigin.z, 2);
                    //MessageBox.Show(imge.UserOrigin.x.ToString() + " , " + imge.UserOrigin.y.ToString() + " , " + imge.UserOrigin.z.ToString());
                    //MessageBox.Show(vVectorArrayIsos[n].x.ToString() + " , " + vVectorArrayIsos[n].y.ToString() + " , " + vVectorArrayIsos[n].z.ToString());
                }

                // Create an array with the coordinates of the isocentes without repetition.
                int aux = 1;
                for (int act = 0; act < vVectorArrayIsoDicomPos.Length; act++)
                {
                    for (int exist = aux; exist < vVectorArrayIsoDicomPos.Length; exist++)
                    {
                        if (vVectorArrayIsos[act].x == vVectorArrayIsos[exist].x & vVectorArrayIsos[act].y == vVectorArrayIsos[exist].y & vVectorArrayIsos[act].z == vVectorArrayIsos[exist].z)
                        {
                            vVectorArrayIsos[exist].x = Double.NaN;
                            vVectorArrayIsos[exist].y = Double.NaN;
                            vVectorArrayIsos[exist].z = Double.NaN;
                        }
                    }
                    aux++;
                }

                // To count the number of isocenters.
                int nIsos = 0;
                for (int n = 0; n < vVectorArrayIsoDicomPos.Length; n++)
                {
                    if (!Double.IsNaN(vVectorArrayIsos[n].x))
                    {
                        nIsos++;
                    }
                }
                return nIsos;
            }
            catch (Exception e)
            {
                throw new Exception("GetNIsos function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get the numbers of beams in the plan.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Number of beams in a plan.
        /// </returns>
        public static int GetNumberOfBeams(PlanSetup plan)
        {
            try
            {
                int i = 0;
                foreach (Beam b in plan.Beams)
                {
                    i++;
                }
                return i;
            }
            catch (Exception e)
            {
                throw new Exception("GetNumberOfBeams function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if some field has a Photon energy.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if some field has a photon energy.
        /// </returns>
        public static bool CheckSomeFieldIsRx(PlanSetup plan)
        {
            try
            {
                string energy;

                foreach (Beam b in plan.Beams)
                {
                    energy = b.EnergyModeDisplayName;
                    //MessageBox.Show(energy);
                    if (StringExtensions.Like(energy, "*X"))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch (Exception e)
            {
                throw new Exception("CheckSomeFieldIsRx function Error: \n" + e.Message);
            }
        }


        /// <summary>
        /// Check if the Fields Setup exist and they are correctly configured. Add messages in warning table and verification table.
        /// </summary>
        /// <remarks>
        /// The Fields Setups are correctly configured if: 
        /// - There are two orthogonal Fields Setups. 
        /// - Their energy is 6 MV. 
        /// - They don't have any linked bolus.
        /// - The field size doesn't irradiate the EPID electronics.
        /// </remarks>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        public static void CheckSetupField(PlanSetup plan)
        {
            try {
                string msg = string.Empty;
                string energy = string.Empty;
                double SetupFieldGantry = double.NaN;
                bool nSetupAP = false, nSetupLL = false;
                bool energyCount = false, bolusCount = false, ortogonalCount = false;
                double collAngle = 0;
                double X1 = 0, X2 = 0, Y2 = 0;
                bool longField = false;
                bool checkSetupFields = true;

                if (CheckPlan.CheckSomeFieldIsRx(plan)) {
                    foreach (Beam b in plan.Beams)
                    {
                        energy = b.EnergyModeDisplayName;
                        if (StringExtensions.Like(energy, "*X"))
                        {
                            bool boolSetupField = b.IsSetupField;
                            if (boolSetupField)
                            {
                                SetupFieldGantry = b.ControlPoints[0].GantryAngle;
                                energy = b.EnergyModeDisplayName;

                                // Check if anterior/posterior and lateral Setup Fields exist
                                if ((SetupFieldGantry == 0 | SetupFieldGantry == 180))
                                {
                                    nSetupAP = true;
                                }

                                if ((SetupFieldGantry == 90 | SetupFieldGantry == 270))
                                {
                                    nSetupLL = true;
                                }

                                // Check 6MV energy
                                if (energy != "6X")
                                {
                                    if (!energyCount)
                                    {
                                        msg += ("          - Setup Fields with energy \u2260 6 MV.\n");
                                        energyCount = true;
                                        checkSetupFields = false;
                                    }
                                }

                                // Check linked bolus
                                if (b.Boluses.Any())
                                {
                                    if (!bolusCount)
                                    {
                                        msg += ("          - Bolus linked to Setup Field.\n");
                                        bolusCount = true;
                                        checkSetupFields = false;
                                    }
                                }

                                // Check field size
                                collAngle = b.ControlPoints[0].CollimatorAngle;
                                Y2 = b.ControlPoints[0].JawPositions.Y2;
                                X1 = b.ControlPoints[0].JawPositions.X1;
                                X2 = b.ControlPoints[0].JawPositions.X2;

                                if (!longField)
                                {
                                    if ((collAngle == 0 && Y2 > 110) || (collAngle == 90 && X2 > 110) || (collAngle == 270 && X1 < -110))
                                    {
                                        msg += ("          - Length of Setup fields is too long, EPID electronics might be irradiated.\n");
                                        longField = true;
                                        checkSetupFields = false;
                                    }
                                }
                            }
                        }
                    }

                    if (nSetupAP == false | nSetupLL == false)
                    {
                        if (ortogonalCount == false)
                        {
                            msg += ("          - Please define anterior/posterior and lateral Setup Fields.\n");
                            ortogonalCount = true;
                            checkSetupFields = false;

                        }
                    }
                    //"Problems with Setup Fields:\n"
                   if (!checkSetupFields)
                    {
                        Script.verificationTable.Rows.Add("Setup Fields", false);
                        Script.warningTable.Rows.Add("Problems with Setup Fields:\n" + msg);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("CheckSetupField function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if each field has the same machine name. Add message in warning table if it's necessary.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        public static void CheckAllFieldsSameMachine(PlanSetup plan)
        {
            try
            {
                string machineName = "";
                int nBeams = 0;
                int nBeamsSameName = 0;
                foreach (Beam b in plan.Beams)
                {
                    nBeams++;
                    if (machineName == "")
                    {
                        machineName = b.TreatmentUnit.Id.ToString();
                        nBeamsSameName++;
                    }
                    else
                    {
                        if (b.TreatmentUnit.Id.ToString() == machineName)
                        {
                            nBeamsSameName++;
                        }
                    }

                }

                if (nBeamsSameName != nBeams)
                {
                    Script.warningTable.Rows.Add("Different machines assigned in the plan.");
                }

            }
            catch (Exception e)
            {
                throw new Exception("CheckAllFieldsSameMachine function Error: \n" + e.Message);
            }
        }      

    }
    #endregion

    #region PLAN PRESCRIPTION FUNCTIONS
    /// <summary>
    /// Plan Presctiption Functions
    /// </summary>
    public class CheckPlanPrescription
    {
        /// <summary>
        /// Get number of planned fractions
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Number of planned fractions.
        /// </returns>
        public static int GetNumberOfFracctions(PlanSetup plan)
        {
            try
            {
                int numberOfFractions =
                    plan.UniqueFractionation.NumberOfFractions.Value; // Number of fractions (planned) 

                if (numberOfFractions > 0)
                    return numberOfFractions;
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception e)
            {
                throw new Exception("GetNumberOfFracctions: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get the planned dose per fraction
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Planned dose per fraction.
        /// </returns>
        public static DoseValue GetDosePerFracction(PlanSetup plan)
        {
            try
            {
                DoseValue dosePerFraction =
                    plan.UniqueFractionation.PrescribedDosePerFraction;

                if (dosePerFraction.Dose > 0 && dosePerFraction.ToString() != "N/A")
                    return dosePerFraction;
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception e)
            {
                throw new Exception("GetDosePerFracction: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get a total planned dose
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Total planned dose.
        /// </returns>
        public static DoseValue GetTotalDose(PlanSetup plan)
        {
            try
            {
                DoseValue totalDose =
                    plan.TotalPrescribedDose;

                if (totalDose.Dose > 0 && totalDose.ToString() != "N/A")
                    return totalDose;
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception e)
            {
                throw new Exception("GetTotalDose: \n" + e.Message);
            }
        }


        /// <summary>
        /// Check if RT Prescription (Physician's intent) exists
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context 
        /// </param>
        /// <returns>
        /// True if RT prescription exists.
        /// </returns>
        public static bool CheckRTPrescriptionExists (PlanSetup plan, LocalScriptContext localContext)
        {
            string planUID = plan.UID;
            SqlDataReader reader;
            SqlConnection conn = localContext.MySQLConnection.Conn;
            try
            {
                string cmdStr = @"SELECT dbo.PrescriptionAnatomyItem.PrescriptionAnatomyItemSer
                           FROM dbo.PrescriptionAnatomy,   
                                 dbo.PrescriptionAnatomyItem,   
                                 dbo.Prescription,   
                                 dbo.PlanSetup,   
                                 dbo.RTPlan  
                           WHERE ( dbo.PrescriptionAnatomy.PrescriptionSer = dbo.Prescription.PrescriptionSer ) and  
                                 ( dbo.PrescriptionAnatomy.PrescriptionAnatomySer = dbo.PrescriptionAnatomyItem.PrescriptionAnatomySer ) and  
                                 ( dbo.PlanSetup.PrescriptionSer = dbo.Prescription.PrescriptionSer ) and  
                                 ( dbo.RTPlan.PlanSetupSer = dbo.PlanSetup.PlanSetupSer ) and  
                                 ( ( dbo.RTPlan.PlanUID = @planUID ) ) ";

                using (SqlCommand command = new SqlCommand(cmdStr, conn))
                {    

                    command.Parameters.AddWithValue("@planUID", planUID);
                    reader = command.ExecuteReader();

                    if (reader.Read() == false)
                    {                        
                        reader.Close();
                        return false;
                    }
                }

                reader.Close();

                return true;
            }
            catch (Exception e)
            {
                throw new Exception("CheckRTPrescriptionExists Error: \n" + e.Message);
            }

        }

        /// <summary>
        /// Get the RT Prescription (Physician's intent).
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <returns>
        /// List with volumes and their prescriptions
        ///     List<Tuple<VolumeId, dose per fraction, total dose>>
        /// </returns>
        public static List<Tuple<string,double,double>> GetRTPrescription (PlanSetup plan, LocalScriptContext localContext)
        {
            //Tuple<VolumeId, dose per fraction, total dose>
            List<Tuple<string, double, double>> physiciansIntentVol = new List<Tuple<string, double, double>>();
            string planUID = plan.UID;
            SqlDataReader reader;
            SqlConnection conn = localContext.MySQLConnection.Conn;
            try
            {
                string cmdStr = @"SELECT dbo.PrescriptionAnatomyItem.ItemValue  
                            FROM dbo.PrescriptionAnatomy,   
                                 dbo.PrescriptionAnatomyItem,   
                                 dbo.Prescription,   
                                 dbo.PlanSetup,   
                                 dbo.RTPlan  
                           WHERE ( dbo.PrescriptionAnatomy.PrescriptionSer = dbo.Prescription.PrescriptionSer ) and  
                                 ( dbo.PrescriptionAnatomy.PrescriptionAnatomySer = dbo.PrescriptionAnatomyItem.PrescriptionAnatomySer ) and  
                                 ( dbo.PlanSetup.PrescriptionSer = dbo.Prescription.PrescriptionSer ) and  
                                 ( dbo.RTPlan.PlanSetupSer = dbo.PlanSetup.PlanSetupSer ) and  
                                 ( dbo.PrescriptionAnatomy.AnatomyRole = 2) and
                                 ( ( dbo.RTPlan.PlanUID = @planUID ) ) ";

                using (SqlCommand command = new SqlCommand(cmdStr, conn))  
                {
                    //int a = 0;
                    int i = 0;
                    string volumeId = string.Empty;
                    string dosePerFraction = string.Empty;
                    string totalDose = string.Empty;

                    command.Parameters.AddWithValue("@planUID", planUID);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        try
                        {                           
                            if (i == 0)
                            {
                                volumeId = reader.GetString(0);
                                //MessageBox.Show("volumeId: " + volumeId);
                                i++;
                            }
                            else if (i == 1)
                            {
                                dosePerFraction = reader.GetString(0);
                                //MessageBox.Show("dosePerFraction: " + dosePerFraction);
                                i++;
                            }
                            else if (i == 2)
                            {
                                totalDose = reader.GetString(0);
                                //MessageBox.Show("totalDose: " + totalDose);
                                physiciansIntentVol.Add(Tuple.Create(volumeId, Convert.ToDouble(dosePerFraction), Convert.ToDouble(totalDose)));
                                i = 0;
                            }
                        }
                        catch (Exception)
                        {                            
                            reader.Close();                           
                        }
                    }
                }

                reader.Close();

                if (physiciansIntentVol.Count() <= 0)
                {
                    physiciansIntentVol = null;
                    //physiciansIntentVol.Add(Tuple.Create(string.Empty, double.NaN, double.NaN));
                }

                return physiciansIntentVol;
            }
            catch (Exception e)
            {                 
                throw new Exception("GetRTPrescription Error: \n" + e.Message);
            }
            
        }

        /// <summary>
        /// Check if Planned Dose Prescription matches RT Prescription (Phycisian's Intent). Add messages in warning table and verification table.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>        /// 
        /// <param name="localContext">
        /// Local context
        /// </param>
        public static void CheckDoseRTPrescription (PlanSetup plan, LocalScriptContext localContext)
        {
            string msg = string.Empty;          
            
            try
            {
                //Tuple<VolumeId, dose per fraction, total dose>
                List<Tuple<string, double, double>> physiciansIntentVol = new List<Tuple<string, double, double>>();

                physiciansIntentVol = CheckPlanPrescription.GetRTPrescription(plan, localContext);
                if (physiciansIntentVol != null)
                {
                    double rtTotalDose = CheckPlanPrescription.GetTotalDose(plan).Dose;
                    double rtDosePerFraction = CheckPlanPrescription.GetDosePerFracction(plan).Dose;

                    physiciansIntentVol.Sort((a, b) => b.Item3.CompareTo(a.Item3));     //Order tuples from high to low total dose.

                    //string msg1 = "rtTotalDose: " + rtTotalDose + "\n";
                    //msg1 += "rtDosePerFraction: " + rtDosePerFraction + "\n";
                    //msg1 += "Physician Dose Per Fraction (physiciansIntentVol[0].Item2): " + physiciansIntentVol[0].Item2 + "\n";
                    //msg1 += "Physician Total Dose(physiciansIntentVol[0].Item3)  : " + physiciansIntentVol[0].Item3 + "\n";
                    //msg1 += "physiciansIntentVol[0].Item3 == rtTotalDose : " + (physiciansIntentVol[0].Item3 == rtTotalDose) + "\n";
                    //msg1 += "resta: " + (physiciansIntentVol[0].Item3 - rtTotalDose) + "\n";
                    //msg1 += "physiciansIntentVol[0].Item2 == rtDosePerFraction : " + (physiciansIntentVol[0].Item2 == rtDosePerFraction) + "\n";

                    //MessageBox.Show(msg1);


                    if (Math.Abs(physiciansIntentVol[0].Item3 - rtTotalDose) > 0.001 ||
                        Math.Abs(physiciansIntentVol[0].Item2 - rtDosePerFraction) > 0.001)
                    {
                        msg += "RT Prescription doesn't match Plan Dose Prescription:\n";
                        msg += "          - RT Prescription: \n";
                        msg += "                Dose Per Fraction:   " + physiciansIntentVol[0].Item2 + " Gy\n";
                        msg += "                Total Dose:               " + physiciansIntentVol[0].Item3 + " Gy\n";
                        msg += "          - RT Plan:\n";
                        msg += "                Dose Per Fraction:   " + rtDosePerFraction + " Gy\n";
                        msg += "                Total Dose:               " + rtTotalDose + " Gy\n";

                        Script.warningTable.Rows.Add(msg);
                        Script.verificationTable.Rows.Add("RT Prescription \u2260 Plan Prescription", false);
                       
                    }
                }
                else
                {                   Script.verificationTable.Rows.Add("RT Prescription exists", false);
                   
                }
            }
            catch (Exception e)
            {
                throw new Exception("CheckDoseRTPrescription Error: \n" + e.Message);
            }

        }
    }
    #endregion

    #region PLAN CALCULATION FUNCTIONS
    /// <summary>
    /// Plan Calculation Functions
    /// </summary>
    public class CheckPlanCalculation
    {             
        /// <summary>
        ///  Get the total number of MU.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Monitor Units
        /// </returns>
        public static double GetMUTotal(PlanSetup plan)
        {
            try
            {
                double muTotal = 0;
                //double dosePerFractionInGy = plan.UniqueFractionation.PrescribedDosePerFraction.Dose;
                foreach (Beam b in plan.Beams)
                {
                    if (Math.Round(b.Meterset.Value, 0) > 0) // MU > 0
                    {
                        muTotal = Math.Round(b.Meterset.Value, 0) + muTotal;
                    }
                }
                return muTotal;
            }
            catch (Exception e)
            {
                throw new Exception("GetMUTotal function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get plan monitor units per Gy
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Plan monitor units per Gy
        /// </returns>
        public static double GetPlanMUperGy(PlanSetup plan)
        {
            try
            {
                double MUperGy = Math.Round((GetMUTotal(plan) / CheckPlanPrescription.GetDosePerFracction(plan).Dose), 2);
                return MUperGy;

            }
            catch (Exception e)
            {
                throw new Exception("GetMUTotal function Error: \n" + e.Message);
            }
        }
    }
    #endregion

    #endregion 

    #region BEAM FUNCTIONS

    /// <summary>
    /// Beam Functions
    /// </summary>
    public class CheckBeam
    {
        /// <summary>
        /// Check if any field without MLC or with static MLC exceeds the minimum units monitors. Add message in warning table.
        /// to advice a rep rate of 600 MU/min. 
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        public static void Check600RepRateAdvice (PlanSetup plan, LocalScriptContext localContext)
        {
            try
            {
                string msg = string.Empty;
                int MUForRR6 = localContext.Calculation.MUForRR6;
                bool advice = false;
                foreach (Beam b in plan.Beams)
                {
                    if (b.Meterset.Value >= MUForRR6 &&
                        (b.MLCPlanType.ToString() == "NotDefined" || b.MLCPlanType.ToString() == "Static")
                && b.DoseRate != 600)
                    {
                        msg += "          - " + b.Id + "\n";
                        advice = true;
                    }
                }
                if (advice)
                {
                    Script.warningTable.Rows.Add("Suggestion: consider using 600 MU/min in field(s) over " + localContext.Calculation.MUForRR6 + " MU: \n" + msg);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Check600RepRateAdvice function Error: \n" + e.Message);
            }
                        
        }     
    }
    #endregion
    
    #region STRUCTURE FUNCTIONS
    /// <summary>
    /// Structure Functions
    /// </summary>
    public class CheckStructure
    {
        /// <summary>
        /// Get the body structure.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// The body structure or null if not exists.
        /// </returns>
        public static Structure GetBodyStructure(PlanSetup plan)
        {
            try
            {
                StructureSet structureSet = plan.StructureSet;
                Structure body = null;
                foreach (var structure in structureSet.Structures)
                {
                    if (structure.DicomType == "EXTERNAL")
                    {
                        body = structure;
                        return body;
                    }
                }
                return body;
            }
            catch (Exception e)
            {
                throw new Exception("GetBodyStructure function Error: \n" + e.Message);
            }

        }

        /// <summary>
        /// Get the Target Structure
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset
        /// </param>
        /// <returns>
        /// Target structure.
        /// </returns>
        public static Structure GetTargetVolumeStructure(PlanSetup plan)
        {
            try
            {
                StructureSet structureSet = plan.StructureSet;
                Structure target = null;
                foreach (var structure in structureSet.Structures)
                {
                    if (structure.Id == plan.TargetVolumeID)
                    {
                        target = structure;
                        break;
                    }
                }

                return target;
            }
            catch (Exception e)
            {
                throw new Exception("GetTargetVolumeStructure function Error: \n" + e.Message);
            }
        }           
        
    }
    #endregion
    
    #region DOSE VOLUME FUNCTIONS
    /// <summary>
    /// Dose Volume Functions
    /// </summary>
    public class CheckDoseVolume
    {
        /// <summary>
        /// Checking the body structure is covered completely by the calculation grid. Add messages in warning table.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>

        public static void CheckBodyDoseCoverage(PlanSetup plan)
        {
            try
            {
                DoseValue DV = new DoseValue(0.0, "Gy");
                Structure body = CheckStructure.GetBodyStructure(plan);
                double bodyDoseCoverage = plan.GetVolumeAtDose(body, DV, VolumePresentation.Relative);
                if (bodyDoseCoverage < 99.9 | bodyDoseCoverage.ToString() == "NeuN")
                {
                    Script.warningTable.Rows.Add("The body is not covered completely by the calculation grid.");
                }
            }
            catch (Exception e)
            {
                throw new Exception("CheckBodyDoseCoverage function Error: \n" + e.Message);
            }
           

        }

        /// <summary>
        /// Check if the mean dose in the target volume is between minimun and maximum percentage values.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Represents a local context.
        /// </param>
        public static void CheckTargetMeanDose(PlanSetup plan, LocalScriptContext localContext)
        {
            try
            {
                double minMeanTargetDose = localContext.Calculation.MinMeanTargetDose;
                double maxMeanTargetDose = localContext.Calculation.MaxMeanTargetDose;
                DoseValue meanTargetDose = CheckDoseVolume.GetMeanTargetDose(plan, "Absolute");
                DoseValue totalPrescribedDose = CheckPlanPrescription.GetTotalDose(plan);
                bool verifTargetMeanDose;
                if ((meanTargetDose.UnitAsString == totalPrescribedDose.UnitAsString) && 
                    (meanTargetDose.Dose > (totalPrescribedDose.Dose * minMeanTargetDose) && meanTargetDose.Dose < (totalPrescribedDose.Dose * maxMeanTargetDose)))
                {
                    verifTargetMeanDose = true;
                } else
                {
                    verifTargetMeanDose = false;
                }

                Script.verificationTable.Rows.Add("Target Mean Dose (" + (localContext.Calculation.MinMeanTargetDose * 100) + "% - " + (localContext.Calculation.MaxMeanTargetDose * 100) + "%)", verifTargetMeanDose);
            }
            catch (Exception e)
            {
                throw new Exception("CheckTargetMeanDose function Error: \n" + e.Message);
            }
        }


        /// <summary>
        /// Get Mean Target Dose.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="presentation">
        /// Dose and volume presentation in DVH (Absolute or Relative)
        /// </param>
        /// <returns> Mean Target Dose </returns>
        public static DoseValue GetMeanTargetDose(PlanSetup plan, string presentation)
        {
            try
            {
                DVHData dvhData = null;
                Structure target = CheckStructure.GetTargetVolumeStructure(plan);
                
                if (String.Equals(presentation, "relative", StringComparison.OrdinalIgnoreCase))
                {
                    dvhData = plan.GetDVHCumulativeData(target,
                        DoseValuePresentation.Relative,
                        VolumePresentation.Relative, 0.1);
                }
                else if (String.Equals(presentation, "absolute", StringComparison.OrdinalIgnoreCase))
                {
                    dvhData = plan.GetDVHCumulativeData(target,
                      DoseValuePresentation.Absolute,
                      VolumePresentation.AbsoluteCm3, 0.1);
                }
                DoseValue meanDose = dvhData.MeanDose;
                return meanDose;
            }
            catch (Exception e)
            {
                throw new Exception("GetMeanTargetDose function Error: \n" + e.Message);
            }

        }

        /// <summary>
        /// Get the dose value which covers a specific percentage of target
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="percent">
        /// Percentage dose (95% -> value: 95 not 0.95)
        /// </param>
        /// <returns>
        /// Dose value.
        /// </returns>
        public static double GetPercentTargetCoverage(PlanSetup plan, double percent)
        {
            try
            {
                Structure target = CheckStructure.GetTargetVolumeStructure(plan);
                DoseValue DV = new DoseValue(percent, "%");
                double V95 = plan.GetVolumeAtDose(target, DV, VolumePresentation.Relative);
                return Math.Round(V95, 1);
            }
            catch (Exception e)
            {
                throw new Exception("GetPercentTargetCoverage function Error: \n" + e.Message);
            }

        }

        /// <summary>
        /// Get the maximum dose (%) in the body
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Dose Value
        /// </returns>
        public static DoseValue GetMaxBodyDose(PlanSetup plan)
        {
            try
            {
                Structure body = CheckStructure.GetBodyStructure(plan);
                DVHData dvhData = plan.GetDVHCumulativeData(body,
                  DoseValuePresentation.Relative,
                  VolumePresentation.Relative, 0.1);

                return dvhData.MaxDose;
            }
            catch (Exception e)
            {
                throw new Exception("GetMaxBodyDose function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get relative dose at volume.
        /// </summary>
        /// <param name="volume">
        /// volume to analize.
        /// </param>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Dose Value or "NeuN" if the volume is not covered completely by the calculation grid.
        /// </returns>

        public static DoseValue GetRelativeDoseAtVolume(double volume, PlanSetup plan, Structure str)
        {
            try
            {
                //Structure body = CheckStructure.GetBodyStructure(plan);
                return plan.GetDoseAtVolume(str, volume, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Relative);
            }
            catch (Exception e)
            {
                throw new Exception("GetRelativeDoseAtVolume function Error: \n" + e.Message);
            }
        }
    }
    #endregion

    #region IMAGE FUNCTIONS
    /// <summary>
    /// Image Functions
    /// </summary>
    public class CheckImage
    {
        /// <summary>
        /// Check if the User Origin coordinates are different to x = 0 mm, y = 0 mm and z = 0 mm. Add messages in Verification table and warning table.
        /// </summary>
        /// <param name="imge">
        /// Represents a CT dataset. 
        /// </param>
        public static void CheckUserOriginModified(Image imge)
        {
            bool verifUserOrigin;            
            try
            {
                if (imge.UserOrigin.x == 0 & imge.UserOrigin.y == 0 & imge.UserOrigin.z == 0)
                {
                    verifUserOrigin = false;
                }
                else
                {
                    verifUserOrigin = true;
                }
                Script.verificationTable.Rows.Add("User Origin \u2260 (0,0,0)", verifUserOrigin);
            }
            catch (Exception e)
            {
                throw new Exception("CheckUserOriginModified function Error: \n" + e.Message);
            }

        }
    }
    #endregion

    #region SQL FUNCTIONS

    /// <summary>
    /// SQL Functions
    /// </summary>
    public class CheckSQL
    {
        /// <summary>
        /// Open a SQL connection
        /// </summary>
        /// <param name="dataSource"> Server name </param>
        /// <param name="initialCatalog"> Database name</param>
        /// <param name="userID"> User name </param>
        /// <param name="password"> Password </param>
        /// <param name="timeOut"> Time(in seconds) property to wait while trying to establish a connection 
        /// before terminating the attempt and generating an error. </param>
        /// <returns>
        /// SQL connection instance.
        /// </returns>
        public static SqlConnection SQLOpenConnection(string dataSource, string initialCatalog, string userID, string password, int timeOut)
        {
            try
            {

                SqlConnection conn = new SqlConnection(); // Open connection to a SQL Server database
                conn.ConnectionString =
                "Data Source="+dataSource+";" +
                "Initial Catalog="+initialCatalog+";" +
                "User id="+userID+";" +
                "Password=" + password + ";" +
                "Connection Timeout=" + timeOut + ";";
                if (conn != null && conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                return conn;
            }
            catch (Exception e)
            {
                throw new Exception("SQLOpenConnection function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Close SQL connection.
        /// </summary>
        /// <param name="conn"> SQL connection instance.</param>
        public static void SQLCloseConnection(SqlConnection conn)
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
            catch (Exception e)
            {
                throw new Exception("SQLCloseConnection function Error: \n" + e.Message);
            }
        }
    }

    #endregion

    #region STRING EXTENSIONS
    /// <summary>
    /// String type management
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Compares the string against a given pattern.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <param name="pattern">The pattern to match, where "*" means any sequence of characters, 
        /// and "?" means any single character.</param>
        /// <returns><c>true</c> if the string matches the given pattern; otherwise <c>false</c>.</returns>
        public static bool Like(this string str, string pattern)
        {
            return new Regex(
                "^" + Regex.Escape(pattern).Replace(@"\*", ".*").Replace(@"\?", ".") + "$",
                RegexOptions.IgnoreCase | RegexOptions.Singleline
            ).IsMatch(str);
        }
    }
    #endregion

    #region PRINTING FUNCTIONS

    /// <summary>
    /// Printing functions
    /// </summary>
    public class CheckPrinting
    {
        //"\u2714" is the "ok" symbol in unicode
        //"\u2718" is the "x" symbol in unicode
        //"\u0020" is the "space" symbol in unicode
        //"\u2022" is the "bullet" symbol in unicode

        /// <summary>
        /// Print the verification Table.
        /// </summary>
        /// <param name="verificationTable">
        /// Verification data table.
        /// </param>
        public static void PrintVerifMessage(DataTable verificationTable)
        {
            try
            {
                if (verificationTable != null && verificationTable.Rows.Count > 0)
                {
                    int dataRowCount = 0;
                    int trueCount = 0;
                    foreach (DataRow row in verificationTable.Rows)
                    {
                        dataRowCount++;
                        if ((bool)row["booleanBool"] == true)
                        {
                            trueCount++;
                        }
                        //MessageBox.Show(row["verificationMessage"] + "\t" + row["booleanBool"].ToString() + "\n");
                    }

                    string printVerificationMessage = "";
                    if (dataRowCount == trueCount)
                    {
                        foreach (DataRow row in verificationTable.Rows)
                        {
                            printVerificationMessage +=
                                ("\u2714     " + row["verificationMessage"] + "\n");  //"\u2714" is the "ok" symbol in unicode
                        }
                    }
                    else
                    {
                        foreach (DataRow row in verificationTable.Rows)
                        {
                            if ((bool)row["booleanBool"] == false)
                            {
                                printVerificationMessage +=
                                    ("\u2718   " + row["verificationMessage"] + "\n"); //"\u2718" is the "x" symbol in unicode
                            }
                            else
                            {
                                printVerificationMessage +=
                                    ("\u0020     " + row["verificationMessage"] + "\n"); //"\u0020" is the "space" symbol in unicode
                            }
                        }
                    }

                    MessageBox.Show(printVerificationMessage, "Verification");
                }
            }
            catch (Exception e)
            {
                throw new Exception("PrintVerifMessage function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Print a warning table.
        /// </summary>
        /// <param name="warningTable">
        /// Warning data table.
        /// </param>
        public static void PrintWarningMessage(DataTable warningTable)

        {
            try
            {
                string printWarningMessage = string.Empty;

                //MessageBox.Show(warningTable.Rows.Count.ToString());

                if (warningTable != null && warningTable.Rows.Count > 0)
                {
                    foreach (DataRow row in warningTable.Rows)
                    {
                        printWarningMessage += ("\u2022    " + row["warningMessage"] + "\n");
                    }

                    MessageBox.Show(printWarningMessage, "Warnings");
                }
                else
                {
                    MessageBox.Show("\n\n No warnings\n\n", "Warnings");
                }
            }
            catch (Exception e)
            {
                throw new Exception("PrintWarningMessage function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Print the information table.
        /// </summary>
        /// <param name="informationTable">
        /// Information data table.
        /// </param>
        public static void PrintInformationMessage(DataTable informationTable)
        {
            try
            {
                string printInformationMessage = string.Empty;

                //MessageBox.Show(informationTable.Rows.Count.ToString());

                if (informationTable != null && informationTable.Rows.Count > 0)
                {
                    foreach (DataRow row in informationTable.Rows)
                    {
                        printInformationMessage += ("    " + row["informationTable"] + "\n");
                    }

                    MessageBox.Show(printInformationMessage, "Information");
                }
                else
                {
                    MessageBox.Show("\n\n No Information \n\n", "Warnings");
                }
            }
            catch (Exception e)
            {
                throw new Exception("PrintInformationMessage function Error: \n" + e.Message);
            }
        }
    }
    #endregion

    // ----------------------------------- End Namespace ----------------------------------------------------
}
#endregion