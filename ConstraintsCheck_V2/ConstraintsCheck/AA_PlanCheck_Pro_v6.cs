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
//using System.Windows.Forms;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

#region NAMESPACE VMS.TPS
namespace VMS.TPS
{    
    #region LOCAL SCRIPT CONTEXT CLASSES
    /// <summary>
    /// Defining the local context.
    /// </summary>
    public class LocalScriptContext
    {
        /// <summary>
        /// Local context constructor.
        /// </summary>
        /// <param name="localContext"> Tuple with classes that are in the local context.</param>
        public LocalScriptContext(Tuple<MachinesList, MySQLConnection, Calculation, Collision> localContext)
        {
            if (localContext.Item1 is MachinesList)
            {
                Machines = localContext.Item1;
            }
            if (localContext.Item2 is MySQLConnection)
            {
                MySQLConnection = localContext.Item2;
            }
            if (localContext.Item3 is Calculation)
            {
                Calculation = localContext.Item3;
            }
            if (localContext.Item4 is Collision)
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

            /// ------ MACHINES ---- 
            var listOfMachine = new List<Machine>();

            // Exact IGRT Couch
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

            // Exact Couch
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

            //Trilogy
            Machine trilogy = new Machine();
            trilogy.MachineId = "Trilogy";
            trilogy.Couch = couchExactIGRT;
            listOfMachine.Add(trilogy);

            //iX
            Machine iX = new Machine();
            iX.MachineId = "iX";
            iX.Couch = couchExactIGRT;
            listOfMachine.Add(iX);

            //URTE
            Machine URTE = new Machine();
            URTE.MachineId = "URTE";
            URTE.Couch = couchExactCouch;
            listOfMachine.Add(URTE);

            //2100CD - gegant
            Machine CD2100 = new Machine();
            CD2100.MachineId = "2100CD";
            CD2100.Couch = couchExactCouch;
            listOfMachine.Add(CD2100);

            Machine[] arrayOfMachine = listOfMachine.ToArray();
            MachinesList machines = new MachinesList(arrayOfMachine);

            /// ------SQLCONNECTION----
            MySQLConnection SQLconn = new MySQLConnection();
            SQLconn.DataSource = "VARIANSRV";
            SQLconn.InitialCatalog = "variansystem";
            SQLconn.UserID = "reports";
            SQLconn.Password = "reports";
            SQLconn.TimeOut = 10;
            SQLconn.Conn = PlanCheckSQL.SQLOpenConnection(SQLconn.DataSource, SQLconn.InitialCatalog, SQLconn.UserID, SQLconn.Password, SQLconn.TimeOut);

            /// ------ CALCULATION ---- 
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

            /// ------ COLLISION ---- 
            Collision collision = new Collision();
            collision.MaxCouchRotCalc = 10; // degrees
            collision.MaxCouchRotWarning = 10; //degrees
            collision.CouchVertPositionCTCorrection = 69.3; //mm
            collision.CollisionSafetyMarginDistance = 10; //mm
            collision.CollisionSafetyMarginGantryAngle = 2; // degrees        

            /// ------ LOCAL CONTEXT GENERATION   -----
            Tuple<MachinesList, MySQLConnection, Calculation, Collision> tuplaObject = new Tuple<MachinesList, MySQLConnection, Calculation, Collision>(machines, SQLconn, calculation, collision);
            LocalScriptContext localContext = new LocalScriptContext(tuplaObject);
            return localContext;
        }
    }
    #endregion

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
            if (PlanCheckVerification.CheckInitialParameters(context, plan, imge, out initialMsg))
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
                MessageBox.Show(e.Message);
            }

            //MachinesList machines = localContext.Machines;
            //Collision coll = localContext.Collision;
            //Calculation calc = localContext.Calculation;
            //MySQLConnection sqlConn = localContext.MySQLConnection;
            //SqlConnection conn = sqlConn.Conn;
            

            /// ----------------- Creating message tables --------------
            // Information messages table
            DataTable informationTable = new DataTable();
            informationTable.Columns.Add("informationTable", typeof(string)); 

            // Verification messages table
            DataTable verificationTable = new DataTable();
            verificationTable.Columns.Add("verificationMessage", typeof(string));
            verificationTable.Columns.Add("booleanBool", typeof(bool));

            // Warning messages table
            DataTable warningTable = new DataTable();
            warningTable.Columns.Add("warningMessage", typeof(string));

            //// ------------------ Information ------------------
            string infMeanTargetDose = null;
            try
            {                
                infMeanTargetDose = PlanCheckDoseVolume.GetMeanTargetDose(plan, "Relative").ToString();
            }
            catch (Exception)
            {
                infMeanTargetDose = ("N/A");
            }

            try 
            {
                informationTable.Rows.Add("Plan :\t" + plan.Id);
                informationTable.Rows.Add();
                informationTable.Rows.Add("Total Dose:\t" + PlanCheckPlanPrescription.GetTotalDose(plan).ToString());
                informationTable.Rows.Add("Dose/fraction :\t" + PlanCheckPlanPrescription.GetDosePerFracction(plan).ToString());
                informationTable.Rows.Add();
                informationTable.Rows.Add("Target volume:\t" + plan.TargetVolumeID);
                informationTable.Rows.Add("     Mean Dose:\t" + infMeanTargetDose);
                informationTable.Rows.Add("     Coverage V95 :\t" + PlanCheckDoseVolume.GetPercentTargetCoverage(plan, 95).ToString() + " %");
                informationTable.Rows.Add();
                informationTable.Rows.Add("Body:");
                informationTable.Rows.Add("     Dmax :\t" + PlanCheckDoseVolume.GetMaxBodyDose(plan));
                informationTable.Rows.Add("     Dmax(2cm3) :\t" + PlanCheckDoseVolume.GetRelativeDoseAtVolume(2.0, plan, PlanCheckStructure.GetBodyStructure(plan)));
                informationTable.Rows.Add();
                informationTable.Rows.Add("Number of Isos :\t" + PlanCheckPlan.GetNIsos(plan, imge));
                informationTable.Rows.Add();
                informationTable.Rows.Add("MU total:\t" + PlanCheckPlanCalculation.GetMUTotal(plan).ToString() + " MU");
                informationTable.Rows.Add("MU/Gy :\t" + Math.Round(PlanCheckPlanCalculation.GetPlanMUperGy(plan), 0).ToString() + " MU/Gy");
            } 
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }


            //------------------ Verification and warnings ------------------        
            
            // Primary Reference Dose and Planned Dose Prescription
            try
            {
                // Check if Planned Dose Prescription percentage is = 100%
                double prescribedPercentage = PlanCheckPlanPrescription.GetPrescribedPercentage(plan) * 100.0;
                if (prescribedPercentage != 100.0)
                {
                    verificationTable.Rows.Add("Prescription Percentage = 100%", false);
                    warningTable.Rows.Add("Prescription Percentage \u2260 100%\n           - " + prescribedPercentage + "%");
                }

                // Check if primary reference point is assigned       
                DoseValue dosePerFractionInPrimaryRefPoint = new DoseValue(0.0, "Gy");
                if (!PlanCheckRefPoints.GetDosePerFractionInPrimaryRefPoint(plan, out dosePerFractionInPrimaryRefPoint))
                {                    
                    warningTable.Rows.Add("Primary Reference Point not assigned.");
                }

                // Check if the dose per fraction in primary reference point is the same as the plan dose.
                bool verifPlanPrescription = PlanCheckRefPoints.CheckDosePlanEqualPriRefPoint(PlanCheckPlanPrescription.GetDosePerFracction(plan), dosePerFractionInPrimaryRefPoint);
                verificationTable.Rows.Add("Dose 100% in Primary Reference Point", verifPlanPrescription);


                //PlanCheckRefPoints.GetDosePerFractionInPrimaryRefPoint(plan, out dosePerFractionInPrimaryRefPoint);
                //DoseValue dosePerFraction =PlanCheckPlanPrescription.GetDosePerFracction(plan);
                //double dosePerFractionDose = dosePerFraction.Dose;

                //MessageBox.Show("dosePerFractionInPrimaryRefPoint: " + dosePerFractionInPrimaryRefPoint + "\n" + 
                //    "dosePerFractionDose: " + dosePerFractionDose);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            
            try { 
                // Check if some reference point has already received dose.
                string msg = string.Empty;
                if (PlanCheckRefPoints.CheckRefPointWithDose(plan, localContext, out msg))
                {
                    warningTable.Rows.Add("Some Reference Points have already received dose:\n" + msg);
                }

                //Check if Planned Dose Prescription matches RT Prescription (Phycisian's Intent).
                msg = string.Empty;
                if (PlanCheckPlanPrescription.CheckRTPrescriptionExists(plan, localContext))
                {
                    if (!PlanCheckPlanPrescription.CheckDoseRTPrescription(plan, localContext, out msg))
                    {
                        verificationTable.Rows.Add("RT Prescription \u2260 Plan Prescription", false);
                        warningTable.Rows.Add("RT Prescription doesn't match Plan Dose Prescription:\n" + msg);
                    }
                }

                // Target mean dose.
                bool verifTargetMeanDose = PlanCheckDoseVolume.CheckTargetMeanDose(plan, localContext);
                verificationTable.Rows.Add("Target Mean Dose (" + (localContext.Calculation.MinMeanTargetDose * 100) + "% - " + (localContext.Calculation.MaxMeanTargetDose * 100) + "%)", verifTargetMeanDose);
                
                // User Origin
                bool verifUserOrigin = PlanCheckImage.CheckUserOriginModified(imge);
                verificationTable.Rows.Add("User Origin \u2260 (0,0,0)", verifUserOrigin);

                // Setup Fields            
                if (PlanCheckPlan.CheckSomeFieldIsRx(plan))
                {
                    msg = string.Empty;
                    bool verifSetupFields = PlanCheckPlan.CheckSetupField(plan, out msg);
                    if (!verifSetupFields)
                    {
                        verificationTable.Rows.Add("Setup Fields", false);
                        warningTable.Rows.Add("Problems with Setup Fields:\n" + msg);
                    }
                }

                // DRRs
                if (PlanCheckPlan.CheckSomeFieldIsRx(plan))
                {                    
                        verificationTable.Rows.Add("DRRs", PlanCheckPlan.CheckDRR(plan));
                }

                // SSD Electron
                if (PlanCheckPlan.CheckSomeFieldIsElectron(plan))
                {
                    verificationTable.Rows.Add("Electron SSD (999 - 1001 mm)", PlanCheckPlan.CheckSSDE(plan));
                }   

                // Cb Name in electron beams
                if (PlanCheckPlan.CheckCbExists(plan) && PlanCheckPlan.CheckSomeFieldIsElectron(plan))
                {
                    bool verifCbName = PlanCheckPlan.CheckCbName(plan);
                    verificationTable.Rows.Add("Cb Name", verifCbName);
                }
            
                // Bolus Names and bolus in each field
                if (PlanCheckPlan.CheckBolusExists(plan))
                {
                    bool verifBolusName = PlanCheckPlan.CheckBolusName(plan);
                    verificationTable.Rows.Add("Bolus Name", verifBolusName);
                    if (!PlanCheckPlan.CheckAllFieldsHaveBolus(plan))
                    {
                        warningTable.Rows.Add("Bolus not linked to all fields or different boluses are used.");
                    }
                }
            
                //Check selected RepRate (according to technique).            
                msg = string.Empty;
                if (!PlanCheckPlan.CheckRepRate(plan, localContext, out msg))
                {
                    verificationTable.Rows.Add("RepRate used", false); //TechniqueRR[0] = VMAT
                    warningTable.Rows.Add(msg);
                }
            
                // Check if all fields with dMLC has 6MV
                msg = string.Empty;
                if (!PlanCheckBeam.CheckDMLCEnergy(plan,out msg))
                {
                    verificationTable.Rows.Add("dMLC fields with 6MV", false);
                    warningTable.Rows.Add("dMLC field without 6MV: \n" + msg);
                } 
                else if (PlanCheckBeam.checkDMLCExists(plan))
                {
                    verificationTable.Rows.Add("dMLC fields with 6MV", true);
                }
            
                // Check half beam jaws.
                string beamMsg = string.Empty;
                string badJawsHalfBeamPosMsg = string.Empty;
                bool badJawsHalfBeamPos = PlanCheckPlan.CheckHalfBeamJaws(plan, out beamMsg, out badJawsHalfBeamPosMsg);

                if (badJawsHalfBeamPos)
                {
                    verificationTable.Rows.Add("Half Beam with Jaws = 0", false);
                    warningTable.Rows.Add(badJawsHalfBeamPosMsg);
                }
                else if (beamMsg != string.Empty)
                {
                    verificationTable.Rows.Add("Half Beam with Jaws = 0", true);
                    warningTable.Rows.Add(beamMsg);
                }

                // Check if delta couch is correct (The plan must be Approved)
                string msgCheckDeltaCouch = String.Empty;
                msg = String.Empty;
                // Status: UnApproved - PlanningApproved - Rejected - TreatmentApproved
                if ((plan.ApprovalStatus.ToString() == "PlanningApproved" || plan.ApprovalStatus.ToString() == "TreatmentApproved") && imge.ImagingOrientation.ToString() == "HeadFirstSupine")
                {
                    bool checkDeltaCouch = PlanCheckPlan.CheckDeltaCouch(plan, imge, localContext, out msg);

                    verificationTable.Rows.Add("Delta Couch Shift", checkDeltaCouch);

                    if (!checkDeltaCouch)
                    {
                        warningTable.Rows.Add("Delta couch values don't match for the following fields:\n" + msg);
                    }
                }
            
                // Warning more than 1 iso
                int infNIsos = PlanCheckPlan.GetNIsos(plan, imge);
                if (infNIsos > 1)
                {
                    warningTable.Rows.Add("Plan with more than 1 isocenter: " + infNIsos + " isos.");
                }

                // --- Check calculation models ---
                // RX
                if (PlanCheckPlan.CheckSomeFieldIsRx(plan))
                {
                    if (!PlanCheckPlanCalculation.CheckPhotonCalcAlgorithm(plan, localContext))
                    {
                        warningTable.Rows.Add("Check the dose calculation algorithm for photons.");
                    }
                }
                // Electrons
                if (PlanCheckPlan.CheckSomeFieldIsElectron(plan))
                {
                    if (!PlanCheckPlanCalculation.CheckElectronCalcAlgorithm(plan, localContext))
                    {
                        warningTable.Rows.Add("Check the dose calculation algorithm for electrons.");
                    }
                }

                // Check if every field has the same machine 
                if (!PlanCheckPlan.CheckAllFieldsSameMachine(plan))
                {
                    warningTable.Rows.Add("Different machines assigned in the plan.");
                }

                //Check if some field has couch rotation over ±(localContext.Collision.MaxCouchRotWarning) degrees 
                if (PlanCheckPlan.CheckCouchRotation(localContext, plan))
                {
                    warningTable.Rows.Add("Couch Rotation > " + localContext.Collision.MaxCouchRotWarning.ToString() + " degrees.");
                }

                // Check if every field has a multiple of 0.5 cm displacement
                if (!PlanCheckPlan.CheckCouchShift(plan, imge))
                {
                    warningTable.Rows.Add("Couch shift not multiple of 0.5cm for some field(s).");
                }
                
                // Check the body structure is covered completely by the calculation grid.
                if (!PlanCheckDoseVolume.CheckBodyDoseCoverage(plan))
                {
                    warningTable.Rows.Add("The body is not covered completely by the calculation grid.");
                }

                // If Trilogy is used the couch must be inserted. 
                if (PlanCheckPlan.CheckSomeTrilogyMachine(plan) && !PlanCheckStructure.CheckInsertCorrectCouch(plan, localContext))
                {                    
                    verificationTable.Rows.Add("Trilogy Couch", false);
                }

                // If couch is inserted than the couch model and HUs must be properly assigned.
                if (PlanCheckStructure.CheckCouchInsert(plan)) // Check if the couch is inserted.
                {
                    if (PlanCheckStructure.CheckInsertCorrectCouch(plan, localContext))
                    {
                        msg = string.Empty;
                        bool couchHuAssigned = PlanCheckStructure.CheckHUCouchAssigned(plan, localContext, out msg);
                        if (!couchHuAssigned)
                        {
                            verificationTable.Rows.Add("Couch HUs assigned", false);
                            warningTable.Rows.Add("Verify the couch HUs assigned: \n" + msg);
                        }
                    }
                    else
                    {
                        verificationTable.Rows.Add("Couch inserted", false);
                        warningTable.Rows.Add("The couch model inserted doesn't correspond to the machine used. \n");
                    }
                }

                //List of volumes with HU assigned not taking account the couch structures.
                msg = string.Empty;
                if (PlanCheckStructure.ListStructHUAssiged(plan, out msg))
                {
                    warningTable.Rows.Add("Volumes with assigned HUs: \n" + msg);
                }

                // Check if it's advisable to use 600UM/min
                msg = string.Empty;
                if (PlanCheckBeam.Check600RepRateAdvice(plan, localContext, out msg))
                {
                    warningTable.Rows.Add("Suggestion: consider using 600 MU/min in field(s) over " + localContext.Calculation.MUForRR6 + " MU: \n" + msg);
                }

                // Check if an extended gantry is used correctly.
                msg = string.Empty;
                if (!PlanCheckBeam.CheckGantryOptimDirecction(plan, imge, out msg, localContext))
                {
                    warningTable.Rows.Add("Check extended gantry option in: \n" + msg);
                }


                // Check if every structure whose name includes the string "DRR" (or similar) is projected into DRRs
                msg = string.Empty;
                if (plan.ApprovalStatus.ToString() == "PlanningApproved" || plan.ApprovalStatus.ToString() == "TreatmentApproved")
                {
                    if (!PlanCheckStructure.CheckStructuresLinkedDRR(plan, localContext, out msg))
                    {
                        warningTable.Rows.Add("Some DRR structures are not projected into DRRs: \n" + msg);
                        verificationTable.Rows.Add("Structures projected into DRRs", false);
                    }
                    else if (PlanCheckStructure.CheckStructuresNames(plan, "*DRR*"))
                    {
                        verificationTable.Rows.Add("Structures projected into DRRs", true);
                    }
                }

                // Check if Tolerance Table is assigned for each field
                msg = string.Empty;
                if (!PlanCheckBeam.CheckToleranceTable(plan, localContext, out msg))
                {
                    verificationTable.Rows.Add("Tolerance table", false);
                    warningTable.Rows.Add("Some field(s) without tolerance table assigned: \n" + msg);
                }



                // -------------------------------------------------
                // ----- Check collision with patient or couch. ----
                // -------------------------------------------------

                // --- Select the couch vert position used for collision prediction.
                double couchVertPosition = double.NaN;
                couchVertPosition = PlanCheckCollisions.GetCouchVertPostion(localContext, imge, plan);


                // --- Calculating collisions and getting the verification and warning messages.           
                if (couchVertPosition != double.NaN) //If the vertical position of the couch is valid, the collision is calculated.
                {

                    Tuple<string, bool, bool> collisionMessage = new Tuple<string, bool, bool>(null, false, false); // Tuple<Warning message Text, exist collision with couch, exist collision with patient>
                    collisionMessage = PlanCheckCollisions.CheckCollisions(plan, localContext, imge, couchVertPosition);    // Check collisions and returns messages and collision results.

                    string warningMessage = collisionMessage.Item1;
                    bool couchCollision = collisionMessage.Item2;
                    bool patientCollision = collisionMessage.Item3;

                    verificationTable.Rows.Add("Collision with couch", couchCollision);
                    verificationTable.Rows.Add("Collision with patient", patientCollision);

                    if (collisionMessage.Item1 != string.Empty)
                    {
                        warningTable.Rows.Add(warningMessage);  // Collision messages.
                    }
                }
                else
                {
                    verificationTable.Rows.Add("Collision with couch", false);
                    verificationTable.Rows.Add("Collision with patient", false);
                    warningTable.Rows.Add("The couch position value is not valid for collisions calculation.");
                }
            // ----- END check collision with patient or couch. ----            
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            // ---------------------------------------------------------------------
            // -----------------------  Printing -----------------------------------
            // ---------------------------------------------------------------------

            // ---------------  Printing verification and warning table ------------
            //"\u2714" is the "ok" symbol in unicode
            //"\u2718" is the "x" symbol in unicode
            //"\u0020" is the "space" symbol in unicode
            //"\u2022" is the "bullet" symbol in unicode

            try
            {
                PlanCheckPrinting.PrintInformationMessage(informationTable);
                PlanCheckPrinting.PrintVerifMessage(verificationTable);
                PlanCheckPrinting.PrintWarningMessage(warningTable);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            // --------- Close SQL connection ---------------
            try
            {
                PlanCheckSQL.SQLCloseConnection(localContext.MySQLConnection.Conn);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

    }

    // ---------------------------------- End Script class ------------------------------------------------
    #endregion
            
    #region PLAN FUNCTIONS

    /// <summary>
    /// General Plan Functions.
    /// </summary>
    public class PlanCheckPlan
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
                VVector[] vVectorArrayIsoDicomPos = new VVector[PlanCheckPlan.GetNumberOfBeams(plan)];
                VVector[] vVectorArrayIsos = new VVector[PlanCheckPlan.GetNumberOfBeams(plan)];
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
        /// Check if some field has an specific treatment technique.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if each field has an Arc Technique.
        /// </returns>
        public static bool CheckTechnique(PlanSetup plan, string technique)
        {
            try
            {
                foreach (Beam b in plan.Beams)
                {
                    if (b.Technique.Id.ToString() == technique)
                    {                        
                        return true;
                    }
                }

                return false;

            }
            catch (Exception e)
            {
                throw new Exception("CheckArcTechnique function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if the plan has some Trilogy machine field.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if some field has Trilogy machine
        /// </returns>
        public static bool CheckSomeTrilogyMachine(PlanSetup plan)
        {
            try
            {
                foreach (Beam b in plan.Beams)
                {
                    if (b.TreatmentUnit.Id.ToString() == "Trilogy")
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                throw new Exception("CheckSomeTrilogyMachine function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if some field has Cerrobend.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if some field has Cerrobend
        /// </returns>
        public static bool CheckCbExists(PlanSetup plan)
        {
            try
            {
                foreach (Beam b in plan.Beams)
                {
                    if (b.Blocks.Any())
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                throw new Exception("CheckCbExists function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if some field has Bolus. 
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if some field has Bolus.
        /// </returns>
        public static bool CheckBolusExists(PlanSetup plan)
        {
            try
            {
                foreach (Beam b in plan.Beams)
                {
                    if (b.Boluses.Any())
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                throw new Exception("CheckBolusExists function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if each field has a photon energy.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if each field has a photon energy.
        /// </returns>
        public static bool CheckEachFieldIsRx(PlanSetup plan)
        {
            try
            {
                string energy;
                foreach (Beam b in plan.Beams)
                {
                    energy = b.EnergyModeDisplayName;
                    //MessageBox.Show(energy);
                    if (!StringExtensions.Like(energy, "*X"))
                    {
                        return false;
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                throw new Exception("CheckEachFieldIsRx function Error: \n" + e.Message);
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
        /// Check if some field has an Electron energy.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if some field has an electron energy.
        /// </returns>
        public static bool CheckSomeFieldIsElectron(PlanSetup plan)
        {
            try
            {
                string energy;

                foreach (Beam b in plan.Beams)
                {
                    energy = b.EnergyModeDisplayName;
                    //MessageBox.Show(energy);
                    if (StringExtensions.Like(energy, "*E"))
                    {
                        return true;
                    }
                }

                return false;

            }
            catch (Exception e)
            {
                throw new Exception("CheckSomeFieldIsElectron function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if the Fields Setup exist and they are correctly configured.
        /// </summary>
        /// <remarks>
        /// The Fields Setups are correctly configured if: 
        /// - There are two orthogonal Fields Setups. 
        /// - Their energy is 6 MV. 
        /// - They don't have any linked bolus.
        /// - The field size doesn't irradiate the EPID electronics.
        /// - The DDR is not assigned.
        /// </remarks>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if the Fields Setup are correctly configured.
        /// Not returns false. If the Fields Setups are not correctly configured sends an exception with a string message.
        /// </returns>
        public static bool CheckSetupField(PlanSetup plan, out string msg)
        {
            try {
                msg = string.Empty;
                string energy = string.Empty;
                double SetupFieldGantry = double.NaN;
                bool nSetupAP = false, nSetupLL = false;
                bool energyCount = false, bolusCount = false, ortogonalCount = false, drr = false; 
                double collAngle = 0;
                double X1 = 0, X2 = 0, Y2 = 0;
                bool longField = false;
                bool checkSetupFields = true;


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

                            // Check drr
                            if (!drr)
                            {
                                try
                                {
                                    string DRRId = b.ReferenceImage.Id;
                                }
                                catch (Exception)
                                {
                                    msg += ("          - DRR is not assigned.\n");
                                    drr = true;
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
                return checkSetupFields;
            }
            catch (Exception e)
            {
                throw new Exception("CheckSetupField function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if each photon field has DRR for static technique.
        /// </summary>
        /// <remarks>
        /// Field Setups are always checked.
        /// </remarks>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if each photon field has DRR.
        /// </returns>
        public static bool CheckDRR(PlanSetup plan)
        {
            try
            {
                string energy;
                foreach (Beam b in plan.Beams)
                {
                    energy = b.EnergyModeDisplayName;
                    if (StringExtensions.Like(energy, "*X") & b.Technique.Id.ToString() != "ARC")
                    {
                        try
                        {
                            string DRRId = b.ReferenceImage.Id;
                        }
                        catch (Exception)
                        {
                            return false;
                        }
                    }

                }
                return true;
            }
            catch (Exception e)
            {
                throw new Exception("CheckDRR function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if each electron field has the SSD between 999 mm and 1001 mm.  
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if the SSD of each electron field is correct.
        /// </returns>
        public static bool CheckSSDE(PlanSetup plan)
        {
            try
            {
                string energy;
                foreach (Beam b in plan.Beams)
                {
                    energy = b.EnergyModeDisplayName;
                    if (StringExtensions.Like(energy, "*E"))
                    {
                        //MessageBox.Show("b.SSD " + b.SSD + " \n");
                        if (b.SSD <= 999 | b.SSD >= 1001)
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                throw new Exception("CheckSSDE function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if the Cb name is correct for each electron beam.
        /// </summary>
        /// <remarks>
        /// Incorrect name: "Block" or "Block" + Number (Example: Block1)
        /// </remarks>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True the Cb name is correct.
        /// </returns>
        public static bool CheckCbName(PlanSetup plan)
        {
            try
            {
                string pattern = @"Block\d{1,2}$";  //"\d{1,2}" -> Match one or two decimal digits. Example: Block1, Block12
                                                    //          -> $ The match must occur at the end of the string or before \n at the end of the line or string. 
                                                    //          -> Pattern: Block1, Block12 // Not Pattern: Block1cm, Block 1cm, Block_1cm
                string input = "";
                string energy;

                foreach (Beam b in plan.Beams)
                {
                    energy = b.EnergyModeDisplayName;
                    if (StringExtensions.Like(energy, "*E"))
                    {
                        if (b.Blocks.Any())
                        {
                            foreach (var block in b.Blocks)
                            {
                                input = block.Id.ToString();
                                MatchCollection matches = Regex.Matches(input, pattern);
                                if (input == "Block")
                                {
                                    return false;
                                }
                                else foreach (Match match in matches)
                                    {
                                        return false;
                                    }
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                throw new Exception("CheckCbName function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if the Bolus name is correct.
        /// </summary>
        /// <remarks>
        ///  Incorrect name: "Bolus" or "Bolus" + Number (Example: Bolus1)
        /// </remarks>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True the bolus name is correct.
        /// </returns>
        public static bool CheckBolusName(PlanSetup plan)
        {
            try
            {
                string pattern = @"Bolus\d{1,2}$";   //"\d{1,2}" -> Match one or two decimal digits. Example: Bolus1, Bolus12
                                                    //          -> $ The match must occur at the end of the string or before \n at the end of the line or string. 
                                                    //          -> Pattern: Bolus1, Bolus12 // Not Pattern: Bolus1cm, Bolus 1cm, Bolus_1cm
                string input = "";


                foreach (Beam b in plan.Beams)
                {
                    if (b.Boluses.Any())
                    {

                        foreach (var bolus in b.Boluses)
                        {
                            input = bolus.Id.ToString();
                            MatchCollection matches = Regex.Matches(input, pattern);
                            if (input == "Bolus")
                            {
                                return false;
                            }
                            else foreach (Match match in matches)
                                {
                                    return false;
                                }
                        }
                    }
                }

                return true;
            }
            catch (Exception e)
            {
                throw new Exception("CheckBolusName function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if each field has a bolus linked and it's the same. 
        /// Fields Setup are not checked.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if each field has bolus and it's the same.
        /// </returns>
        public static bool CheckAllFieldsHaveBolus(PlanSetup plan)
        {
            try
            {
                int nBolus = 0;
                int nBeams = 0;
                bool severalBolus = false;
                string refBolusID = "";

                foreach (Beam b in plan.Beams)
                {
                    if (!b.IsSetupField) // It's not a Field Setup.
                    {
                        nBeams++;

                        if (b.Boluses.Any())
                        {
                            foreach (var bolus in b.Boluses)
                            {
                                if (nBolus == 0)
                                {
                                    refBolusID = bolus.Id.ToString();
                                }
                                else
                                {
                                    if (refBolusID != bolus.Id.ToString())
                                    {
                                        severalBolus = true;
                                    }
                                }
                                nBolus++;
                            }
                        }
                    }
                }
                if (nBeams == nBolus && severalBolus == false)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                throw new Exception("CheckAllFieldsHaveBolus function Error: \n" + e.Message);
            }
        }        

        /// <summary>
        /// Check if every beam with treatment Technique parameter has a dose rate especified.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext"></param>
        /// <param name="msg"></param>
        /// <returns></returns>
        public static bool CheckRepRate(PlanSetup plan, LocalScriptContext localContext, out string msg)
        {
            try
            {                
                msg = string.Empty;
                bool checkRepRate = true;
                bool first = true;
                bool sameTechnique = true;
                string auxTechinque = string.Empty;
                foreach (Beam b in plan.Beams)
                {
                    string bTechnique = PlanCheckPlan.GetTechnique(b, localContext);
                    Tuple<string, int> contextTechnique = localContext.Calculation.TechniqueRR.Find(item => item.Item1.Equals(bTechnique));
                    
                    try // This "try" is necessary to avoid the error "Object reference not set to instance of an object" if the bTecnica does not match any TechniqueRR.
                    {
                        if (bTechnique == contextTechnique.Item1 && b.DoseRate != contextTechnique.Item2)
                        {
                            if (bTechnique != auxTechinque)
                            {
                                sameTechnique = false;
                            }
                            if (first)
                            {
                                msg += ("Some " + contextTechnique.Item1 + " field(s) with RepRate \u2260 " + contextTechnique.Item2 + " MU/min: \n");
                                auxTechinque = bTechnique;
                                sameTechnique = true;
                                first = false;
                            } else if (!sameTechnique) {
                                msg += ("       Some " + contextTechnique.Item1 + " field(s) with RepRate \u2260 " + contextTechnique.Item2 + " MU/min: \n");
                                auxTechinque = bTechnique;
                                sameTechnique = true;
                                first = false;
                            }
                            msg += "          - " + b.Id + "\n";
                            checkRepRate = false;
                        }
                    }
                    catch { }
                }
                return checkRepRate;
            }
            catch (Exception e)
            {
                throw new Exception("CheckRepRate function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get Treatment technique for a field.
        /// </summary>
        /// <param name="b">Treatment field</param>
        /// <param name="localContext"> Local Context</param>
        /// <returns>Technique name. Empty if it doesn't exist.</returns>
        public static string GetTechnique(Beam b, LocalScriptContext localContext)
        {
            string technique = string.Empty;
            if (b.Technique.Id.ToString() == "ARC" && b.MLCPlanType.ToString() == "VMAT")
            {
                technique = localContext.Calculation.Techniques[0]; // "VMAT";
            } 
            else if (b.Technique.Id.ToString() == "ARC" && (b.MLCPlanType.ToString() == "ArcDynamic" || b.MLCPlanType.ToString() == "NotDefined"))
            {
                technique = localContext.Calculation.Techniques[1]; // "ARC";
            }
             else if (b.Technique.Id.ToString() == "STATIC" && b.MLCPlanType.ToString() == "DoseDynamic")
            {
                technique = localContext.Calculation.Techniques[2]; // "IMRT";
            } 
            else if (b.Technique.Id.ToString() == "STATIC" && (b.MLCPlanType.ToString() == "Static" || b.MLCPlanType.ToString() == "NotDefined"))
            {
                technique = localContext.Calculation.Techniques[3]; // "3D";
            }

            return technique;
        }

        /// <summary>
        /// Check Half Beam Jaws
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// String message with the result of the check.
        /// </returns>
        public static bool CheckHalfBeamJaws(PlanSetup plan, out string beamMsg, out string badJawsHalfBeamPosMsg)
        {

            int numSup = 0;
            int numInf = 0;
            bool problem = false;
            int numFullBeams = 0;

            beamMsg = string.Empty;
            badJawsHalfBeamPosMsg = ("Please verify the following jaws for Half Beams:\n");
            bool badJawsHalfBeamPosBool = false;

            try
            {
                foreach (Beam b in plan.Beams)
                {
                    bool boolSetupField = b.IsSetupField;
                    if (!boolSetupField)
                    {
                        // Collimator angle 0 degrees. Jaw position in mm
                        if (b.ControlPoints[0].CollimatorAngle == 0 & ((Math.Abs(b.ControlPoints[0].JawPositions.Y1) <= 5) | (Math.Abs(b.ControlPoints[0].JawPositions.Y2) <= 5)))
                        {
                            if (Math.Abs(b.ControlPoints[0].JawPositions.Y1) <= 5 & b.ControlPoints[0].JawPositions.Y2 > 10) { numSup++; }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.Y2) <= 5 & b.ControlPoints[0].JawPositions.Y1 < -10) { numInf++; }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.Y1) <= 5 & Math.Abs(b.ControlPoints[0].JawPositions.Y1) > 0.01)
                            {
                                badJawsHalfBeamPosMsg += ("          - " + b.Id + " with Y1 = " + (-1) * b.ControlPoints[0].JawPositions.Y1 / 10 + "\n"); // (-1) scale correction.
                                problem = true;

                            }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.Y2) <= 5 & Math.Abs(b.ControlPoints[0].JawPositions.Y2) > 0.01)
                            {
                                badJawsHalfBeamPosMsg += ("          - " + b.Id + " with Y2 = " + b.ControlPoints[0].JawPositions.Y2 / 10 + "\n");
                                problem = true;

                            }
                        }

                        if (b.ControlPoints[0].CollimatorAngle == 0 & Math.Abs(b.ControlPoints[0].JawPositions.Y1) >= 5 & Math.Abs(b.ControlPoints[0].JawPositions.Y2) >= 5)
                        {
                            beamMsg += ("          - " + b.Id + "\n");
                            numFullBeams++;
                        }

                        // Collimator angle 90 degrees. Jaw position in mm
                        if (b.ControlPoints[0].CollimatorAngle == 90 & ((Math.Abs(b.ControlPoints[0].JawPositions.X1) <= 5) | (Math.Abs(b.ControlPoints[0].JawPositions.X2) <= 5)))
                        {
                            if (Math.Abs(b.ControlPoints[0].JawPositions.X1) <= 5 & b.ControlPoints[0].JawPositions.X2 > 10) { numSup++; }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.X2) <= 5 & b.ControlPoints[0].JawPositions.X1 < -10) { numInf++; }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.X1) <= 5 & Math.Abs(b.ControlPoints[0].JawPositions.X1) > 0.01)
                            {
                                badJawsHalfBeamPosMsg += ("          - " + b.Id + " with X1 = " + (-1) * b.ControlPoints[0].JawPositions.X1 / 10 + "\n"); // (-1) scale correction.
                                problem = true;

                            }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.X2) <= 5 & Math.Abs(b.ControlPoints[0].JawPositions.X2) > 0.01)
                            {
                                badJawsHalfBeamPosMsg += ("          - " + b.Id + " with X2 = " + b.ControlPoints[0].JawPositions.X2 / 10 + "\n");
                                problem = true;

                            }
                        }

                        if (b.ControlPoints[0].CollimatorAngle == 90 & Math.Abs(b.ControlPoints[0].JawPositions.X1) >= 5 & Math.Abs(b.ControlPoints[0].JawPositions.X2) >= 5)
                        {
                            beamMsg += ("          - " + b.Id + "\n");
                            numFullBeams++;
                        }

                        // Collimator angle 270 degrees. Jaw position in mm
                        if (b.ControlPoints[0].CollimatorAngle == 270 & ((Math.Abs(b.ControlPoints[0].JawPositions.X1) <= 5) | (Math.Abs(b.ControlPoints[0].JawPositions.X2) <= 5)))
                        {
                            if (Math.Abs(b.ControlPoints[0].JawPositions.X1) <= 5 & b.ControlPoints[0].JawPositions.X2 > 10) { numInf++; }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.X2) <= 5 & b.ControlPoints[0].JawPositions.X1 < -10) { numSup++; }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.X1) <= 5 & Math.Abs(b.ControlPoints[0].JawPositions.X1) > 0.01)
                            {
                                badJawsHalfBeamPosMsg += ("          - " + b.Id + " with X1 = " + (-1) * b.ControlPoints[0].JawPositions.X1 / 10 + "\n"); // (-1) scale correction.
                                problem = true;

                            }
                            if (Math.Abs(b.ControlPoints[0].JawPositions.X2) <= 5 & Math.Abs(b.ControlPoints[0].JawPositions.X2) > 0.01)
                            {
                                badJawsHalfBeamPosMsg += ("          - " + b.Id + " with X2 = " + b.ControlPoints[0].JawPositions.X2 / 10 + "\n");
                                problem = true;

                            }
                        }

                        if (b.ControlPoints[0].CollimatorAngle == 270 & Math.Abs(b.ControlPoints[0].JawPositions.X1) >= 5 & Math.Abs(b.ControlPoints[0].JawPositions.X2) >= 5)
                        {
                            beamMsg += ("          - " + b.Id + "\n");
                            numFullBeams++;
                        }
                    }
                }

                // MessageBox.Show("numSup: " + numSup + "\n " + "numInf: " + numInf + "\n " + "problem: " + problem.ToString() + "\n ");
                if (numSup > 0 & numInf > 0 & problem == true)
                {

                    if (numFullBeams == 1)
                    {
                        badJawsHalfBeamPosMsg = (numFullBeams + " beam in plan is NOT half beam:\n") + beamMsg + "\n      " + badJawsHalfBeamPosMsg;
                    }
                    else if (numFullBeams > 1)
                    {
                        badJawsHalfBeamPosMsg = (numFullBeams + " beams in plan are NOT half beams:\n") + beamMsg + "\n       " + badJawsHalfBeamPosMsg;
                    }
                    // badJawsHalfBeamPosMsg = ("Please verify the following jaws for Half Beams:\n") + badJawsHalfBeamPosMsg; // When the plan has half beams and some of them have a bad jaw position.


                    badJawsHalfBeamPosBool = true;
                }


                if (numSup > 0 & numInf > 0 & numFullBeams > 0)
                {

                    if (numFullBeams == 1)
                    {
                        beamMsg = (numFullBeams + " beam in plan is NOT half beam:\n") + beamMsg;

                    }
                    else if (numFullBeams > 1)
                    {
                        beamMsg = (numFullBeams + " beams in plan are NOT half beams:\n") + beamMsg;
                    }
                }
                else if (numSup > 0 & numInf > 0)
                {
                    beamMsg = ("Half Beams are used");
                }
                else
                {
                    //MessageBox.Show("numSup: " + numSup + "\n " + "numInf: " + numInf + "\n " + "problem: " + problem.ToString() + "\n ");
                    beamMsg = string.Empty; // Half-Beams are not present.
                }

                return badJawsHalfBeamPosBool;
            }
            catch (Exception e)
            {
                throw new Exception("  CheckHalfBeamJaws2 function Error: \n" + e.Message);
            }
        }        

        /// <summary>
        /// Check if each field has the same machine name. 
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True each field has the same machine name.
        /// </returns>
        public static bool CheckAllFieldsSameMachine(PlanSetup plan)
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

                if (nBeamsSameName == nBeams)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                throw new Exception("CheckAllFieldsSameMachine function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if "delta couch" matches with the isocenter coordinates from the user origin (displacements).
        /// </summary>
        /// <param name="plan">Represents a treatment plan dataset.</param>
        /// <param name="imge">Represents a CT dataset. </param>
        /// <param name="localContext">Local Context</param>
        /// <param name="msg">Beams where delta couch does't match</param>
        /// <returns>false if delta couch doesn`t match</returns>
        public static bool CheckDeltaCouch(PlanSetup plan, Image imge, LocalScriptContext localContext, out string msg)
        {
            try
            {

                bool deltaCouch = true;
                SqlDataReader reader;
                StructureSet structureSet = plan.StructureSet;
                string planUID = plan.UID;
                string fieldId = "";
                SqlConnection conn = localContext.MySQLConnection.Conn;
                msg = String.Empty;

                foreach (Beam b in plan.Beams)
                {

                    double deltaCouchVrt = double.NaN;
                    double deltaCouchLng = double.NaN;
                    double deltaCouchLat = double.NaN;
                    double isoX = double.NaN, isoY = double.NaN, isoZ = double.NaN;
                    VVector isob = PlanCheckBeam.GetIsocenter(b, imge, plan);

                    // Get Delta couch
                    string cmdStr = @"SELECT dbo.ExternalFieldCommon.CouchLatDelta, dbo.ExternalFieldCommon.CouchLngDelta,dbo.ExternalFieldCommon.CouchVrtDelta  
                                            FROM dbo.RTPlan, dbo.PlanSetup, dbo.Radiation, dbo.ExternalFieldCommon  
                                            WHERE   ( dbo.PlanSetup.PlanSetupSer = dbo.RTPlan.PlanSetupSer ) and  
                                                    ( dbo.Radiation.PlanSetupSer = dbo.PlanSetup.PlanSetupSer ) and  
                                                    ( dbo.ExternalFieldCommon.RadiationSer = dbo.Radiation.RadiationSer ) and  
                                                    ( ( dbo.RTPlan.PlanUID =  @planUID ) AND  
                                                    ( dbo.Radiation.RadiationId = @fieldId))";

                    // MovieGenreID = reader.IsDBNull(movieGenreIDIndex) ? null : reader.GetInt32(movieGenreIDIndex)

                    fieldId = b.Id;
                    using (SqlCommand command = new SqlCommand(cmdStr, conn))
                    {
                        command.Parameters.AddWithValue("@planUID", planUID);
                        command.Parameters.AddWithValue("@fieldId", fieldId);
                        reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            deltaCouchLat = reader.IsDBNull(0) ? double.NaN : ((-1) * (reader.GetDouble(0)) * 10);
                            deltaCouchLng = reader.IsDBNull(1) ? double.NaN : ((-1) * (reader.GetDouble(1)) * 10);
                            deltaCouchVrt = reader.IsDBNull(2) ? double.NaN : ((reader.GetDouble(2)) * 10);

                            if (deltaCouchLat.ToString() == "NeuN" || deltaCouchLng.ToString() == "NeuN" || deltaCouchVrt.ToString() == "NeuN")
                            {
                                reader.Close();
                                return false;
                            }
                        }

                    }
                    reader.Close();

                    isoX = isob.x - imge.UserOrigin.x; // X respect user origin not from dicom. (mm)
                    isoZ = isob.z - imge.UserOrigin.z; // Z respect user origin not from dicom. (mm)
                    isoY = isob.y - imge.UserOrigin.y; // Y respect user origin not from dicom. (mm)  

                    if (Math.Abs(deltaCouchLat - isoX) > 1 || Math.Abs(deltaCouchLng - isoZ) > 1 || Math.Abs(deltaCouchVrt - isoY) > 1) //False if the difference between displacement and Delta Couch is greater 1 mm
                    {
                        msg += ("          - " + b.Id + "\n");
                        deltaCouch = false;
                    }
                }

                return deltaCouch;

            }
            catch (Exception e)
            {
                throw new Exception("CheckDeltaCouch function Error: \n" + e.Message);
            }

        }

        /// <summary>
        /// Check if couch shift is multiple of 5 mm
        /// </summary>
        /// <param name="plan"></param>
        /// Represent a treatment plan dataset.
        /// <param name="imge"></param>
        /// Represents a CT dataset.
        /// <returns>
        /// True if couch shift is multiple of 5 mm
        /// </returns>
        public static bool CheckCouchShift(PlanSetup plan, Image imge)
        {
            try
            {
                String energy;
                double multiple = 5;
                double x, y, z;
                foreach (Beam b in plan.Beams)
                {
                    energy = b.EnergyModeDisplayName;

                    if (StringExtensions.Like(energy, "*X"))
                    {
                        x = b.IsocenterPosition.x - imge.UserOrigin.x;
                        y = b.IsocenterPosition.y - imge.UserOrigin.y;
                        z = b.IsocenterPosition.z - imge.UserOrigin.z;

                        if ((Math.Round(x, 1) % multiple != 0) || (Math.Round(y, 1) % multiple != 0) || (Math.Round(z, 1) % multiple != 0))
                        {
                            return false;
                        }
                    }
                }

                return true;
            }
            catch (Exception e)
            {
                throw new Exception("CheckCouchShift function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if a couch doesn't exceed a degree value.
        /// (Ex. 10 degrees means:  10 < couchAngle > 350)
        /// </summary>
        /// <param name="localContext">Local Context</param>
        /// <param name="plan">Represents a treatment plan dataset.</param>
        /// <returns>true if exceed the degree values.</returns>
        static public bool CheckCouchRotation(LocalScriptContext localContext, PlanSetup plan)
        {
            bool exceedRotation = false; // Not exceed the maximum rotation. 
            double degrees = localContext.Collision.MaxCouchRotWarning;
            foreach (Beam b in plan.Beams)
            {
                if (b.ControlPoints[0].PatientSupportAngle > degrees && b.ControlPoints[0].PatientSupportAngle < (360 - degrees))
                {
                    exceedRotation = true;
                    return exceedRotation;
                }

            }
            return exceedRotation;
        }

    }
    #endregion

    #region PLAN PRESCRIPTION FUNCTIONS
    /// <summary>
    /// Plan Presctiption Functions
    /// </summary>
    public class PlanCheckPlanPrescription
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
        /// Get the prescribed dose percentage as a decimal number
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Dose percentage as a decimal number
        /// </returns>
        public static double GetPrescribedPercentage(PlanSetup plan)
        {
            try
            {
                double prescribedPercentage = plan.PrescribedPercentage; //The prescribed dose percentage as a decimal number
                return prescribedPercentage;
            }
            catch (Exception e)
            {
                throw new Exception("GetPrescribedPercentage: \n" + e.Message);
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
        public static bool CheckRTPrescriptionExists(PlanSetup plan, LocalScriptContext localContext)
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
                    physiciansIntentVol.Add(Tuple.Create(string.Empty, double.NaN, double.NaN));
                }

                return physiciansIntentVol;
            }
            catch (Exception e)
            {                 
                throw new Exception("GetRTPrescription Error: \n" + e.Message);
            }
            
        }

        /// <summary>
        /// Check if Planned Dose Prescription matches RT Prescription (Phycisian's Intent) 
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>        
        /// <param name="localContext">
        /// Local context
        /// </param>
        /// <param name="msg">
        /// Out string parameter with the Planned Dose Prescription and RT Prescription.
        /// </param>
        /// <returns>
        /// True if Planned Dose Prescription matches RT Prescription 
        /// </returns>
        public static bool CheckDoseRTPrescription(PlanSetup plan, LocalScriptContext localContext, out string msg)
        {
            msg = string.Empty;

            try
            {
                //Tuple<VolumeId, dose per fraction, total dose>
                List<Tuple<string, double, double>> physiciansIntentVol = new List<Tuple<string, double, double>>();
                physiciansIntentVol = PlanCheckPlanPrescription.GetRTPrescription(plan, localContext);

                double rtTotalDose = PlanCheckPlanPrescription.GetTotalDose(plan).Dose;
                double rtDosePerFraction = PlanCheckPlanPrescription.GetDosePerFracction(plan).Dose;

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
                    msg += "          - RT Prescription: \n";
                    msg += "                Dose Per Fraction:   " + physiciansIntentVol[0].Item2 + " Gy\n";
                    msg += "                Total Dose:               " + physiciansIntentVol[0].Item3 + " Gy\n";
                    msg += "          - RT Plan:\n";
                    msg += "                Dose Per Fraction:   " + rtDosePerFraction + " Gy\n";
                    msg += "                Total Dose:               " + rtTotalDose + " Gy\n";

                    return false;
                }


                return true;
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
    public class PlanCheckPlanCalculation
    {

        /// <summary>
        /// Check if the Photon Calculation Algorithm is correctly assigned.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// true if the Photon Calculation Algorithm is correctly assigned.
        /// </returns>
        public static bool CheckPhotonCalcAlgorithm(PlanSetup plan, LocalScriptContext localContext)
        {
            try
            {
                string photonCalcModel = localContext.Calculation.PhotonVolumeDoseAlgorithm;
                if (plan.PhotonCalculationModel != photonCalcModel)
                {
                    return false;
                } else
                {
                    return true;
                }
            } catch (Exception e)
            {
                throw new Exception("CheckPhotonCalcAlgorithm function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if the Electron Calculation Algorithm is correctly assigned.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// true if the Electron Calculation Algorithm is correctly assigned.
        /// </returns>
        public static bool CheckElectronCalcAlgorithm(PlanSetup plan, LocalScriptContext localContext)
        {
            try
            {
                string electronCalcModel = localContext.Calculation.ElectronVolumeDoseAlgorithm;
                if (plan.ElectronCalculationModel != electronCalcModel)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception e)
            {
                throw new Exception("CheckElectronCalcAlgorithm function Error: \n" + e.Message);
            }
        }
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
        /// 
        /// </summary>
        /// <param name="plan"></param>
        /// <returns></returns>
        public static double GetPlanMUperGy(PlanSetup plan)
        {
            try
            {
                double MUperGy = Math.Round((GetMUTotal(plan) / PlanCheckPlanPrescription.GetDosePerFracction(plan).Dose), 2);
                return MUperGy;

            }
            catch (Exception e)
            {
                throw new Exception("GetMUTotal function Error: \n" + e.Message);
            }
        }
    }
    #endregion

    #region PLAN REFERENCE POINT FUNCTIONS
    /// <summary>
    /// Referenct point functions.
    /// </summary>
    public class PlanCheckRefPoints
    {
        /// <summary>
        ///  Check if the dose per fraction in primary reference point is the same as the planned dose.
        /// </summary>
        /// <param name="prescribedDosePerFraction">
        /// Represents the planned dose.
        /// </param>
        /// <param name="dosePerFractionInPrimaryRefPoint">
        /// Represents the dose per fraction in primary reference point.
        /// </param>
        /// <returns>
        /// True the dose is the same.
        /// </returns>
        public static bool CheckDosePlanEqualPriRefPoint(DoseValue prescribedDosePerFraction, DoseValue dosePerFractionInPrimaryRefPoint)
        {
            try
            {
                //Check de units and value are the same for Primary reference dose and plan prescribed dose (both per fraction)
                if ((dosePerFractionInPrimaryRefPoint.UnitAsString == prescribedDosePerFraction.UnitAsString) && (dosePerFractionInPrimaryRefPoint.Dose == prescribedDosePerFraction.Dose))
                {
                    return true;
                }
                return false;
            }
            catch (Exception e)
            {
                throw new Exception("CheckDosePlanEqualPriRefPoint function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Checks if some reference point has already received dose.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>        
        /// <param name="localContext">
        /// Local context
        /// </param>
        /// <param name="msg">
        /// Message with the reference points which have already received some dose.
        /// </param>
        /// <returns>
        /// true if some reference point has already received dose.
        /// </returns>
        public static bool CheckRefPointWithDose(PlanSetup plan, LocalScriptContext localContext, out string msg)
        {
            try
            {
                SqlDataReader reader;
                //string patientId = patient.Id;
                string courseId = plan.Course.Id;
                string planSetupId = plan.Id;
				string planUID = plan.UID;
				string patientId = string.Empty;
                msg = string.Empty;                
                bool deliveredDose = false;
                double totDeliveredDose = double.NaN;
                SqlConnection conn = localContext.MySQLConnection.Conn;
				
				
				string cmdStr1 = @"SELECT DISTINCT dbo.Patient.PatientId  
								FROM dbo.Patient,   
									 dbo.Course,   
									 dbo.PlanSetup,   
									 dbo.RTPlan  
								WHERE ( dbo.Course.PatientSer = dbo.Patient.PatientSer ) and  
									  ( dbo.PlanSetup.CourseSer = dbo.Course.CourseSer ) and  
									  ( dbo.RTPlan.PlanSetupSer = dbo.PlanSetup.PlanSetupSer ) and  
									  ( ( dbo.Course.CourseId = @courseId ) AND  
									  ( dbo.RTPlan.PlanUID = @planUID) )"; 
				
				using (SqlCommand command = new SqlCommand(cmdStr1, conn))
                {
                    
                    command.Parameters.AddWithValue("@courseId", courseId);
                    command.Parameters.AddWithValue("@planUID", planUID);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
						patientId = reader.GetString(0);
					}
				}
				reader.Close();
                string cmdStr = @"SELECT DISTINCT dbo.RefPointDeliveredDose.TotDeliveredDose,
                                dbo.RefPoint.RefPointId   
                                FROM dbo.Patient,   
                                     dbo.Course,   
                                     dbo.PlanSetup,   
                                     dbo.RefPoint,   
                                     dbo.RefPointDeliveredDose,   
                                     dbo.RadiationRefPoint,   
                                     dbo.RTPlan  
                               WHERE ( dbo.Course.PatientSer = dbo.Patient.PatientSer ) and  
                                     ( dbo.PlanSetup.CourseSer = dbo.Course.CourseSer ) and  
                                     ( dbo.RefPoint.RefPointSer = dbo.RefPointDeliveredDose.RefPointSer ) and  
                                     ( dbo.RefPoint.PatientSer = dbo.Patient.PatientSer ) and  
                                     ( dbo.RTPlan.RTPlanSer = dbo.RadiationRefPoint.RTPlanSer ) and  
                                     ( dbo.PlanSetup.PlanSetupSer = dbo.RTPlan.PlanSetupSer ) and  
                                     ( dbo.RefPoint.RefPointSer = dbo.RadiationRefPoint.RefPointSer ) and  
                                     ((dbo.Patient.PatientId = @patientId ) AND  
                                     ( dbo.Course.CourseId = @courseId ) AND
                                     ( dbo.PlanSetup.PlanSetupId = @planSetupId))";


                using (SqlCommand command = new SqlCommand(cmdStr, conn))
                {
                    command.Parameters.AddWithValue("@patientId", patientId);
                    command.Parameters.AddWithValue("@courseId", courseId);
                    command.Parameters.AddWithValue("@planSetupId", planSetupId);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        totDeliveredDose = reader.GetDouble(0);
                        if (totDeliveredDose > 0)
                        {
                            msg += ("          - " + reader.GetString(1) + ": " + Math.Round(totDeliveredDose,3) + " Gy\n");
                            deliveredDose = true;
                        }
                        
                    }
                   // MessageBox.Show(msg);
                }
                reader.Close();

                return deliveredDose;
            }
            catch (Exception e)
            {
                throw new Exception("CheckRefPointWithDose function Error: \n" + e.Message);
            }
        }

        
        /// <summary>
        /// Get the assigned dose per fraction in primary reference point
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="dosePerFractionInPrimaryRefPoint">
        /// Out parameter which has the assigned dose per fracction in the primary reference point 
        /// </param>
        /// <returns>
        /// True if the dose per fraction in primary reference point is assigned.
        /// </returns>
        public static bool GetDosePerFractionInPrimaryRefPoint(PlanSetup plan, out DoseValue dosePerFractionInPrimaryRefPoint)
        {
            try
            {
                dosePerFractionInPrimaryRefPoint = new DoseValue(0.0, "Gy");
                if (plan.UniqueFractionation.DosePerFractionInPrimaryRefPoint.Dose.ToString() != "NeuN")
                {
                    dosePerFractionInPrimaryRefPoint = plan.UniqueFractionation.DosePerFractionInPrimaryRefPoint;                    
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                throw new Exception("GetDosePerFractionInPrimaryRefPoint function Error: \n" + e.Message);
            }
        }

    }
    #endregion

    #region BEAM FUNCTIONS

    /// <summary>
    /// Beam Functions
    /// </summary>
    public class PlanCheckBeam
    {
        /// <summary>
        /// Get isocenter corrected by image(CT) and patient(plan) orientation.
        /// </summary>
        /// <param name="b">Beam ID</param>
        /// <param name="imge"> Represents a CT dataset.</param>
        /// <param name="plan">Represents a treatment plan dataset.</param>
        /// <returns> VVector with the isocenter coordinates</returns>
        public static VVector GetIsocenter(Beam b, Image imge, PlanSetup plan)
        {
            //----- PatientOrientation -----
            //NoOrientation             0 No orientation.       ->      NOT USED
            //HeadFirstSupine           1 Head first -supine.
            //HeadFirstProne            2 Head first -prone.
            //HeadFirstDecubitusRight   3 Head first -decubitus right.
            //HeadFirstDecubitusLeft    4 Head first -decubitus left.
            //FeetFirstSupine           5 Feet first -supine.
            //FeetFirstProne            6 Feet first -prone.
            //FeetFirstDecubitusRight   7 Feet first -decubitus right.
            //FeetFirstDecubitusLeft    8 Feet first -decubitus left.
            //Sitting                   9 Sitting.              ->      NOT USED

            double isoX, isoY, isoZ, isoXAux, isoYAux;
            VVector isocenter = new VVector();

            isoX = b.IsocenterPosition.x;
            isoY = b.IsocenterPosition.y;
            isoZ = b.IsocenterPosition.z;
            isoXAux = isoX;
            isoYAux = isoY;

            //string msg2 = "";
            //msg2 += ("iso PRE \n");
            //msg2 += ("isoX: " + Math.Round(isoX, 2) + "    ");
            //msg2 += ("isoY: " + Math.Round(isoY, 2) + "    ");
            //msg2 += ("isoZ: " + Math.Round(isoZ, 2) + "    ");
            //MessageBox.Show(msg2);            

            switch (plan.TreatmentOrientation.ToString())
            {
                //          X   Y   Z
                case "HeadFirstSupine":             //hfs = 	1	7	5  <- Reference signs values
                    isoX = isoX * (1);
                    isoY = isoY * (1);
                    isoZ = isoZ * (1);
                    break;
                case "HeadFirstProne":              //hfp =		-1	-7	5
                    isoX = isoX * (-1);
                    isoY = isoY * (-1);
                    isoZ = isoZ * (1);
                    break;
                case "HeadFirstDecubitusRight":     //hfd =		7	-1	5
                    isoX = isoYAux * (1);
                    isoY = isoXAux * (-1);
                    isoZ = isoZ * (1);
                    break;
                case "HeadFirstDecubitusLeft":      //hfl =		-7	1	5
                    isoX = isoYAux * (-1);
                    isoY = isoXAux * (1);
                    isoZ = isoZ * (1);
                    break;
                case "FeetFirstSupine":             //ffs =		-1	7	-5
                    isoX = isoX * (-1);
                    isoY = isoY * (1);
                    isoZ = isoZ * (-1);
                    break;
                case "FeetFirstProne":              //ffp =		1	-7	-5
                    isoX = isoX * (1);
                    isoY = isoY * (-1);
                    isoZ = isoZ * (-1);
                    break;
                case "FeetFirstDecubitusRight":     //ffr =		-7	-1	-5
                    isoX = isoYAux * (-1);
                    isoY = isoXAux * (-1);
                    isoZ = isoZ * (-1);
                    break;
                case "FeetFirstDecubitusLeft":      //ffl =		7	1	-5  
                    isoX = isoYAux * (1);
                    isoY = isoXAux * (1);
                    isoZ = isoZ * (-1);
                    break;
            }

            isocenter.x = isoX;
            isocenter.y = isoY;
            isocenter.z = isoZ;

            return isocenter;
        }

        /// <summary>
        /// Get the TH value -> vertical distance between couch and isocenter.
        /// </summary>
        /// <param name="couchVertPosition">
        /// couch vertical position relative to the dicom origin (0,0,0)
        /// </param>
        /// <param name="isocenter">
        /// isocenter position relative to the dicom origin (0,0,0)
        /// </param>
        /// <returns>TH value</returns>
        public static double GetTh(double couchVertPosition, VVector isocenter)
        {
            try
            {
                double th;
                th = couchVertPosition - isocenter.y;
                return th;

            }
            catch (Exception e)
            {
                throw new Exception("GetTh function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get the extended range for Varian Auto Sequencie value.
        /// </summary>
        /// <param name="b">Beam ID</param>
        /// <param name="plan">Represents a treatment plan dataset</param>
        /// <param name="conn">Open connection to a SQL Server database</param>
        /// <returns>The possible values are: 
        ///     - NN
        ///     - NE
        ///     - EN
        ///     - EE
        ///     The two characters define the Start and Stop Angles.
        /// </returns>
        public static string GetExtendedFieldValue(Beam b, PlanSetup plan, SqlConnection conn)
        {
            string extended = null;
            string planUID = plan.UID;
            string fieldId = b.Id;
            SqlDataReader reader;
            try
            {
                // Get the extended range for Varian Auto Sequencie value.
                string cmdStr = @"SELECT dbo.ExternalField.GantryRtnExt FROM dbo.RTPlan, 
                                            dbo.PlanSetup,   
                                            dbo.Radiation,   
                                            dbo.ExternalFieldCommon,   
                                            dbo.ExternalField  
                                    WHERE ( dbo.PlanSetup.PlanSetupSer = dbo.RTPlan.PlanSetupSer ) and  
                                            ( dbo.Radiation.PlanSetupSer = dbo.PlanSetup.PlanSetupSer ) and  
                                            ( dbo.ExternalFieldCommon.RadiationSer = dbo.Radiation.RadiationSer ) and  
                                            ( dbo.ExternalField.RadiationSer = dbo.ExternalFieldCommon.RadiationSer ) and  
                                            ( dbo.RTPlan.PlanUID = @planUID ) and
                                            ( dbo.Radiation.RadiationId = @fieldId) ";

                using (SqlCommand command = new SqlCommand(cmdStr, conn))
                {
                    command.Parameters.AddWithValue("@planUID", planUID);
                    command.Parameters.AddWithValue("@fieldId", fieldId);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        extended = (reader.GetString(0));
                    }
                }
                reader.Close();
                return extended;
            }
            catch (Exception e)
            {
                throw new Exception("GetExtendedFieldValue function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if any field without MLC or with static MLC exceeds the minimum units monitors 
        /// to advice a rep rate of 600 MU/min. 
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <param name="msg">
        /// Out string parameter with the field names where 600 MU/min is advisable.
        /// </param>
        /// <returns>
        /// True if 600 MU/min is advisable.
        /// </returns>
        public static bool Check600RepRateAdvice(PlanSetup plan, LocalScriptContext localContext, out string msg)
        {
            msg = string.Empty;
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
            return advice;

        }    


        /// <summary>
        /// Check if some field in a plan has a dynamic MLC assigned.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True if the dynamic MLC is assigned
        /// </returns>
        public static bool checkDMLCExists(PlanSetup plan)
        {
            try
            {
                foreach (Beam b in plan.Beams)
                {
                    if ((b.MLCPlanType.ToString() == "ArcDynamic" || b.MLCPlanType.ToString() == "VMAT" || b.MLCPlanType.ToString() == "DoseDynamic"))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                throw new Exception("checkDMLCExists function Error: \n" + e.Message);
            }

        }


        /// <summary>
        /// Check if every field with dynamic MLC has 6MV assigned.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="msg">
        /// Message with field names which don't have 6MV assigned.
        /// </param>
        /// <returns>
        /// false if some field with dynamic MLC doesn't have 6MV assigned.
        /// </returns>
        public static bool CheckDMLCEnergy (PlanSetup plan, out string msg)
        {
            try {

                msg = string.Empty;
                bool checkDMLC = true;
                foreach (Beam b in plan.Beams)
                {                   
                    if ((b.MLCPlanType.ToString() == "ArcDynamic" || b.MLCPlanType.ToString() == "VMAT" || b.MLCPlanType.ToString() == "DoseDynamic") && b.EnergyModeDisplayName!="6X")
                    {
                        msg += ("          - "+b.Id+"\n");
                        checkDMLC = false;
                    }
                }
                return checkDMLC;
            }
            catch (Exception e)
            {
                throw new Exception("CheckDMLCEnergy function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check for each field if the direction of gantry rotation is optimal.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="imge">
        /// Represents a CT dataset. 
        /// </param>
        /// <param name="msg">
        /// Message with fields where direction of gantry rotation is not optimal.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <returns>
        /// False if some field has not the optimal gantry rotation direction.
        /// </returns>
        public static bool CheckGantryOptimDirecction(PlanSetup plan, Image imge, out string msg, LocalScriptContext localContext)
        {
            try
            {
                bool CouchRotOutOfRange = false;
                bool[] gantryRotSector = new bool[2]; // [0] = Gantry Left (359.9-180.1); [1] = Gantry Right (0-180); 
                string msgNotCalc = string.Empty;
                double startGantryAngle = double.NaN;
                double endGantryAngle = double.NaN;
                int nControlPoints;
                int nSubField;
                msg = string.Empty;
                bool optim = true;
                SqlConnection conn = localContext.MySQLConnection.Conn;

                foreach (Beam b in plan.Beams) // collision for each field.
                {

                    if (localContext.Collision.MaxCouchRotCalc >= (b.ControlPoints[0].PatientSupportAngle) || 360 - localContext.Collision.MaxCouchRotCalc <= (b.ControlPoints[0].PatientSupportAngle))
                    {
                        VVector isocenter = PlanCheckBeam.GetIsocenter(b, imge, plan); // Get isocenter corrected by image(CT) and patient(plan) orientation.

                        if (b.Technique.Id.ToString() == "ARC")
                        {
                            nSubField = 2;
                            startGantryAngle = b.ControlPoints[0].GantryAngle;  // Start Gantry Angle (degrees)
                            nControlPoints = b.ControlPoints.Count();
                            endGantryAngle = b.ControlPoints[nControlPoints - 1].GantryAngle; // End Gantry Angle (degrees)
                        }
                        else
                        {
                            nSubField = 1;
                            startGantryAngle = b.ControlPoints[0].GantryAngle;  // Start Gantry Angle (degrees)
                        }

                        for (int i = 0; i < nSubField; i++)
                        {
                            if (i == 0) // Start Gantry Angle.
                            {
                                gantryRotSector = PlanCheckCollisions.GetRotationSector(plan, b, conn, startGantryAngle); // bool -> [0] = Gantry Left (359.9-180.1); [1] = Gantry Right (0-180)
                                if (isocenter.x > 20 && startGantryAngle > 175 && startGantryAngle < 185 && gantryRotSector[0] == true)
                                {
                                    msg += "          - " + b.Id + "(Start G)\n";
                                    optim = false;

                                }
                                else if (isocenter.x < -20 && startGantryAngle > 175 && startGantryAngle < 185 && gantryRotSector[1] == true)
                                {

                                    msg += "          - " + b.Id + "(Start G)\n";
                                    optim = false;
                                }
                            }
                            else if (i == 1) // End Gantry Angle.
                            {
                                gantryRotSector = PlanCheckCollisions.GetRotationSector(plan, b, conn, endGantryAngle); // bool -> [0] = Gantry Left (359.9-180.1); [1] = Gantry Right (0-180)
                                if (isocenter.x > 20 && endGantryAngle > 175 && endGantryAngle < 185 && gantryRotSector[0] == true)
                                {
                                    msg += "          - " + b.Id + "(End G)\n";
                                    optim = false;
                                }
                                else if (isocenter.x < -20 && endGantryAngle > 175 && endGantryAngle < 185 && gantryRotSector[1] == true)
                                {
                                    msg += "          - " + b.Id + "(End G)\n";
                                    optim = false;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (CouchRotOutOfRange == false)
                        {
                            msgNotCalc += ("          -   Rotation Gantry direc (Couch Rotation > " + localContext.Collision.MaxCouchRotCalc + " degrees): \n");
                            CouchRotOutOfRange = true;
                        }
                        msgNotCalc += ("\t" + b.Id.ToString() + "\n");
                    }


                }

                msg = msg + msgNotCalc;

                return optim;

            }
            catch (Exception e)
            {
                throw new Exception("CheckGantryOptimDirecction function Error: \n" + e.Message);
            }

        }

        /// <summary>
        /// Check if every field has the tolerance table assigned.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <param name="msg">
        /// Message with fields which don't have the tolerance table assigned.
        /// </param>
        /// <returns>
        /// False if the tolerance table is not assigned.
        /// </returns>
        public static bool CheckToleranceTable(PlanSetup plan, LocalScriptContext localContext, out string msg){
            try
            {

                bool TolTableInsert = true;
                SqlDataReader reader;                
                string planUID = plan.UID;
                string fieldId = "";
                SqlConnection conn = localContext.MySQLConnection.Conn;
                msg = String.Empty;

                foreach (Beam b in plan.Beams)
                {
                    string TolTableId = string.Empty;
                    // Get tolerance table ID
                    string cmdStr = @"  SELECT dbo.Tolerance.ToleranceId  
                                                FROM dbo.PlanSetup,   
                                                     dbo.RTPlan,   
                                                     dbo.Tolerance,   
                                                     dbo.ExternalFieldCommon,   
                                                     dbo.Radiation  
                                               WHERE ( dbo.PlanSetup.PlanSetupSer = dbo.RTPlan.PlanSetupSer ) and  
                                                     ( dbo.ExternalFieldCommon.ToleranceSer = dbo.Tolerance.ToleranceSer ) and  
                                                     ( dbo.Radiation.RadiationSer = dbo.ExternalFieldCommon.RadiationSer ) and  
                                                     ( dbo.PlanSetup.PlanSetupSer = dbo.Radiation.PlanSetupSer ) and  
                                                     ( dbo.ExternalFieldCommon.RadiationSer = dbo.Radiation.RadiationSer ) and  
                                                     ( ( dbo.RTPlan.PlanUID =  @planUID ) AND  
                                                     ( dbo.Radiation.RadiationId = @fieldId))";
                   

                    fieldId = b.Id;
                    using (SqlCommand command = new SqlCommand(cmdStr, conn))
                    {
                        
                        command.Parameters.AddWithValue("@planUID", planUID);
                        command.Parameters.AddWithValue("@fieldId", fieldId);
                        reader = command.ExecuteReader();
                      
                        while (reader.Read())
                        {
                            TolTableId = reader.GetString(0);
                        }

                        if (TolTableId == string.Empty)
                        {
                            msg += ("          - " + b.Id + "\n");
                            TolTableInsert = false;
                        }

                    }
                    reader.Close();                   
                }

                return TolTableInsert;

            }
            catch (Exception e)
            {
                throw new Exception("CheckDeltaCouch function Error: \n" + e.Message);
            }
        }
    }
    #endregion

    #region STRUCTURE FUNCTIONS
    /// <summary>
    /// 
    /// </summary>
    public class PlanCheckStructure
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

            //if (body == null)
            //{
            //    throw new ApplicationException("There isn't any Body Structure.");
            //}
        }


        /// <summary>
        /// Get a list with all couch structures in the Structure Set.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// List with all couch structures in the Structure Set.
        /// </returns>
        public static List<Structure> GetPlanCouchStruture(PlanSetup plan)
        {
            try
            {
                StructureSet structureSet = plan.StructureSet;
                List<Structure> couchStructure = new List<Structure>();
                if (structureSet == null)
                    throw new ApplicationException("The selected plan doesn't have a StructureSet.");

                foreach (var structure in structureSet.Structures)
                {
                    if (structure.DicomType == "SUPPORT")
                    {
                        couchStructure.Add(structure);
                    }
                }

                return couchStructure;

            }
            catch (Exception e)
            {
                throw new Exception("GetPlanCouchStruture function Error: \n" + e.Message);
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

        /// <summary>
        /// Check if all couch structures have the HUs correctly assigned.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset 
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <param name="msg">
        /// Message with the fields which have not the HUs correctly assigned.
        /// </param>
        /// <returns>
        /// False if some couch structure has not the HUs correctly assigned
        /// </returns>
        public static bool CheckHUCouchAssigned(PlanSetup plan, LocalScriptContext localContext, out string msg)
        {
            try
            {
                bool assigned = true;
                msg = string.Empty;
                List<Structure> planCouchStructureList = new List<Structure>();
                planCouchStructureList = PlanCheckStructure.GetPlanCouchStruture(plan); // get the couch structures in plan
                String planMachineId = plan.Beams.First().TreatmentUnit.Id; // 

                foreach (Machine m in localContext.Machines)
                {
                    //m.Couch.CouchRegions[0].CouchParts[0].CouchPieceId;
                    //m.Couch.CouchRegions[0].CouchParts[0].HU;
                    if (m.MachineId == planMachineId)
                    {
                        for (int i = 0; i < planCouchStructureList.Count(); i++)
                        {
                            for (int h = 0; h < m.Couch.CouchRegions.Count(); h++)
                            {                               
                                if (m.Couch.CouchRegions[h].CouchRegionName == planCouchStructureList[i].Name)
                                {
                                   
                                    for (int j = 0; j < m.Couch.CouchRegions[h].CouchParts.Count(); j++)
                                    {
                                        double HUAssigned; 
                                        bool HUIsAssigned = planCouchStructureList[i].GetAssignedHU(out HUAssigned);
                                        if (HUIsAssigned && m.Couch.CouchRegions[h].CouchParts[j].CouchPieceId == planCouchStructureList[i].Id && Math.Round(m.Couch.CouchRegions[h].CouchParts[j].HU, 1) != Math.Round(HUAssigned, 1))
                                        {
                                            msg += "          - " + planCouchStructureList[i].Id + ": " + HUAssigned + "HU\n";
                                            assigned = false;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                return assigned;

            }
            catch (Exception e)
            {
                throw new Exception("CheckHUCouchAssigned function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if the couch is inserted.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True the couch is inserted.
        /// </returns>
        public static bool CheckCouchInsert(PlanSetup plan)
        {
            try
            {
                StructureSet structureSet = plan.StructureSet;

                foreach (Beam b in plan.Beams)
                {
                    if (structureSet == null)
                        throw new ApplicationException("The selected plan doesn't have a StructureSet.");

                    foreach (var structure in structureSet.Structures)
                    {
                        if (structure.DicomType == "SUPPORT")
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                throw new Exception("CheckCouchExists function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check if the couch inserted matches with machine in use
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <returns>
        /// true if the couch inserted matches with machine in use.
        /// </returns>
        public static bool CheckInsertCorrectCouch(PlanSetup plan, LocalScriptContext localContext)
        {
            try
            {
                bool inserted = false;                
                List<Structure> planCouchStructureList = new List<Structure>();
                planCouchStructureList = PlanCheckStructure.GetPlanCouchStruture(plan); // get the couch structures in plan
                String planMachineId = plan.Beams.First().TreatmentUnit.Id; // 

                foreach (Machine m in localContext.Machines)
                {
                    //m.Couch.CouchRegions[0].CouchParts[0].CouchPieceId;
                    //m.Couch.CouchRegions[0].CouchParts[0].HU;
                    if (m.MachineId == planMachineId)
                    {
                        for (int i = 0; i < planCouchStructureList.Count(); i++)
                        {
                            for (int h = 0; h < m.Couch.CouchRegions.Count(); h++)
                            {                                
                                if (m.Couch.CouchRegions[h].CouchRegionName == planCouchStructureList[i].Name)
                                {                                   
                                    inserted = true;
                                }
                            }
                        }
                    }
                }
                return inserted;

            }
            catch (Exception e)
            {
                throw new Exception("CheckInsertCorrectCouch function Error: \n" + e.Message);
            }
        }
        

        /// <summary>
        /// List with HU assigned not taking into account the couch structures.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="msg">
        /// Message with the structures which have HU assigned.
        /// </param>
        /// <returns>
        /// False if no structure has HU assigned.
        /// </returns>
        public static bool ListStructHUAssiged(PlanSetup plan, out string msg)
        {
            try
            {
                StructureSet structureSet = plan.StructureSet;
                double huValue;
                msg = string.Empty;
                bool assigned = false;
                bool HUIsAssigned = false;
                foreach (var structure in structureSet.Structures)
                {
                    HUIsAssigned = structure.GetAssignedHU(out huValue);
                    if (HUIsAssigned && structure.DicomType != "SUPPORT")

                    {                        
                        msg += "          - " + structure.Id + ":       " + Math.Round(huValue,0) + " HU.\n";
                        assigned = true;
                    }
                }
                return assigned;
            }
            catch (Exception e)
            {
                throw new Exception("ListStructHUAssiged function Error: \n" + e.Message);
            }
        }


        /// <summary>
        /// Check if a structure name matches with a specific character chain.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="chainName">
        /// specific character chain (Ex. "*DRR*" or "left*") where "*" means any sequence of characters, 
        /// and "?" means any single character
        /// </param>
        /// <returns>
        /// True if structure name matches with the specific character chain.
        /// </returns>
        public static bool CheckStructuresNames(PlanSetup plan, string chainName)
        {
            StructureSet structureSet = plan.StructureSet;
            bool nameStructureExists = false;

            if (structureSet == null)
                throw new ApplicationException("The selected plan doesn't have a StructureSet.");


            foreach (var structure in structureSet.Structures)
            {
                if (StringExtensions.Like(structure.Id, chainName))
                {
                    nameStructureExists = true;
                }
            }

            return nameStructureExists;
        }

        /// <summary>
        /// Check if every structure whose name includes the string "DRR" (or similar) is projected into DRRs
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <param name="msg">
        /// Message with structures whose name includes the string "DRR" and they are not projected into DRRs.
        /// </param>
        /// <returns>
        /// False if some structure whose name includes the string "DRR" is not projected into DRRs.
        /// </returns>
        public static bool CheckStructuresLinkedDRR(PlanSetup plan, LocalScriptContext localContext, out string msg)
        {
            try
            {               
                StructureSet structureSet = plan.StructureSet;               
                List<string> linkedDRRId = new List<string>();
                linkedDRRId = PlanCheckStructure.GetStructuresLinkedDRR(plan, localContext);
                bool linkedEveryStruct = true;
                msg = string.Empty;
                


               if (structureSet == null)
                    throw new ApplicationException("The selected plan doesn't have a StructureSet.");

                foreach (var structure in structureSet.Structures)
                {
                    //if (PlanCheckStructure.CheckStructuresNames(plan, "*DRR*") || PlanCheckStructure.CheckStructuresNames(plan, "*DDR*") || PlanCheckStructure.CheckStructuresNames(plan, "*RDR*") || PlanCheckStructure.CheckStructuresNames(plan, "*DRD*")) 
                    if (StringExtensions.Like(structure.Id, "*DRR*") || StringExtensions.Like(structure.Id, "*DDR*") || StringExtensions.Like(structure.Id, "*RDR*") || StringExtensions.Like(structure.Id, "*DRD*"))
                    {
                        if (!linkedDRRId.Exists(element => element == structure.Id))
                        {
                            msg += "          - " + structure.Id + "\n";
                            linkedEveryStruct = false;
                        }                        
                    }
                }

                return linkedEveryStruct;

            }
            catch (Exception e)
            {
                throw new Exception("CheckStructuresLinkedDRR function Error: \n" + e.Message);
            }


        }

        /// <summary>
        /// Get the Id of structures outlined in reference images (DRRs).
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <returns>
        /// List with the Id of structures outlined in reference images (DRRs)
        /// </returns>
        public static List<string> GetStructuresLinkedDRR(PlanSetup plan, LocalScriptContext localContext)
        {
            try
            {
                SqlDataReader reader;
                string planUID = plan.UID;
                List<string> GraphicAnnotationId = new List<string>();
                string msg = string.Empty;
                SqlConnection conn = localContext.MySQLConnection.Conn;

                // Get GraphicAnnotationId
                string cmdStr = @"SELECT DISTINCT dbo.GraphicAnnotation.GraphicAnnotationId  
                            FROM dbo.GraphicAnnotation,   
                                    dbo.GraphicAnnotationType,   
                                    dbo.Image,   
                                    dbo.PlanSetup,   
                                    dbo.Radiation,   
                                    dbo.RTPlan  
                            WHERE ( dbo.GraphicAnnotationType.GraphicAnnotationTypeSer = dbo.GraphicAnnotation.GraphicAnnotationTypeSer ) and  
                                    ( dbo.Image.ImageSer = dbo.GraphicAnnotation.ImageSer ) and  
                                    ( dbo.Radiation.PlanSetupSer = dbo.PlanSetup.PlanSetupSer ) and  
                                    ( dbo.Radiation.RefImageSer = dbo.Image.ImageSer ) and  
                                    ( dbo.RTPlan.PlanSetupSer = dbo.PlanSetup.PlanSetupSer ) and  
                                    ( ( dbo.GraphicAnnotationType.GraphicAnnotationTypeIndex = 10 ) AND  
                                    ( dbo.RTPlan.PlanUID = @planUID ) ) ";


                using (SqlCommand command = new SqlCommand(cmdStr, conn))
                {
                    command.Parameters.AddWithValue("@planUID", planUID);

                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        GraphicAnnotationId.Add(reader.GetString(0));
                    }
                }
                reader.Close();


                return GraphicAnnotationId;
            }
            catch (Exception e)
            {

                throw new Exception("GetStructuresLinkedDRR Error: \n" + e.Message);
            }
        }
    }
    #endregion

    #region DOSE VOLUME FUNCTIONS
    /// <summary>
    /// 
    /// </summary>
    public class PlanCheckDoseVolume
    {
        /// <summary>
        /// Checking the body structure is covered completely by the calculation grid.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// True the body is covered completely.
        /// </returns>
        public static bool CheckBodyDoseCoverage(PlanSetup plan)
        {

            DoseValue DV = new DoseValue(0.0, "Gy");
            Structure body = PlanCheckStructure.GetBodyStructure(plan);
            double bodyDoseCoverage = plan.GetVolumeAtDose(body, DV, VolumePresentation.Relative);
            if (bodyDoseCoverage < 99.9 | bodyDoseCoverage.ToString() == "NeuN")
            {
                return false;
            }
            return true;

        }

        /// <summary>
        /// Check if the mean dose in the target volume is between localContext parameters.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>        
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <returns>
        /// True - the mean dose is between the specified values.
        /// </returns>
        public static bool CheckTargetMeanDose(PlanSetup plan, LocalScriptContext localContext)
        {
            try
            {
                double minMeanTargetDose = localContext.Calculation.MinMeanTargetDose;
                double maxMeanTargetDose = localContext.Calculation.MaxMeanTargetDose;
                DoseValue meanTargetDose = PlanCheckDoseVolume.GetMeanTargetDose(plan, "Absolute");
                DoseValue totalPrescribedDose = PlanCheckPlanPrescription.GetTotalDose(plan);
                if ((meanTargetDose.UnitAsString == totalPrescribedDose.UnitAsString) &&
                    (meanTargetDose.Dose > (totalPrescribedDose.Dose * minMeanTargetDose) && meanTargetDose.Dose < (totalPrescribedDose.Dose * maxMeanTargetDose)))
                {
                    return true;
                }
                return false;
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
                Structure target = PlanCheckStructure.GetTargetVolumeStructure(plan);
                
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
                throw new Exception("GetMeanTargetDose function Error: \n       Possible Target Volume not Valid \n" + e.Message);
            }

        }

        /// <summary>
        /// Get the dose value which covers a specific percentage of target
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Dose value.
        /// </returns>
        public static double GetPercentTargetCoverage(PlanSetup plan,double percent)
        {
            try
            {
                Structure target = PlanCheckStructure.GetTargetVolumeStructure(plan);
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
        /// <remarks>
        /// The global variable "body" is used.
        /// </remarks>
        public static DoseValue GetMaxBodyDose(PlanSetup plan)
        {
            try
            {
                Structure body = PlanCheckStructure.GetBodyStructure(plan);
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
                //Structure body = PlanCheckStructure.GetBodyStructure(plan);
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
    /// 
    /// </summary>
    public class PlanCheckImage
    {
        /// <summary>
        /// Check if the User Origin coordinates are different to x = 0 mm, y = 0 mm and z = 0 mm
        /// </summary>
        /// <param name="imge">
        /// Represents a CT dataset. 
        /// </param>
        /// <returns>
        /// True if the coordinates of User Origin has been modified.
        /// </returns>
        public static bool CheckUserOriginModified(Image imge)
        {
            try
            {
                if (imge.UserOrigin.x == 0 & imge.UserOrigin.y == 0 & imge.UserOrigin.z == 0)
                {
                    return false;
                }
                return true;
            }
            catch (Exception e)
            {
                throw new Exception("CheckUserOriginModified function Error: \n" + e.Message);
            }

        }
    }
    #endregion

    #region SET OF FUNCTIONS TO CALCULATE GANTRY COLLISION.

    /// <summary>
    /// Set of functions to calculate gantry collision.
    /// </summary>
    public class PlanCheckCollisions
    {
        /// <summary>
        /// Get the couch vertical position. If Couch structure exists the vertical position is calculated from this
        /// but if not it is calculated from CT DICOM position.
        /// </summary>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <param name="imge">
        /// Represents a CT dataset.
        /// </param>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <returns>
        /// Couch vertical position in mm
        /// </returns>
        public static double GetCouchVertPostion(LocalScriptContext localContext, Image imge, PlanSetup plan)
        {
            try
            {
                double couchVertPositionCT = double.NaN;
                double couchVertPositionStructure = double.NaN;
                double couchVertPosition = double.NaN;

                SqlConnection conn = localContext.MySQLConnection.Conn;

                // Couch vert position from CT.
                try
                {
                    couchVertPositionCT = PlanCheckCollisions.GetCouchVertPositionCT(localContext, imge);
                }
                catch (Exception) // Example, when the CT doesn't exist (phamtom image) 
                {
                    couchVertPositionCT = double.NaN;
                }

                // Couch vert porsition from couch structure
                try
                {
                    couchVertPositionStructure = PlanCheckCollisions.GetCouchVertPositionStructure(plan, localContext);
                }
                catch (Exception)
                {
                    couchVertPositionStructure = double.NaN;
                }

                // Select the couch vert position to calculate the collisions.                
                if (couchVertPositionCT.ToString() != "NeuN")
                {
                    couchVertPosition = couchVertPositionCT;
                }
                else if (couchVertPositionStructure.ToString() != "NeuN")
                {
                    couchVertPosition = couchVertPositionStructure;
                }

                return couchVertPosition;
            }
            catch (Exception e)
            {
                throw new Exception("GetCouchVertPostion Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get the couch vertical position relative to the dicom origin (0,0,0) calculated from the CT.
        /// </summary>
        /// <param name="localContext"> 
        /// Local Context 
        /// </param>
        /// <param name="imge"> 
        /// Represents a CT dataset.
        /// </param>
        /// <returns>
        /// Couch vertical position in mm
        /// </returns>
        public static double GetCouchVertPositionCT(LocalScriptContext localContext, Image imge)
        {
            SqlDataReader reader;
            try
            {
                string UIDSerie = imge.Series.UID;
                double couchVertPositionCT = double.NaN;
                SqlConnection conn = localContext.MySQLConnection.Conn;

                string cmdStr = "SELECT DISTINCT dbo.Slice.CouchVrt FROM dbo.Series, dbo.Slice WHERE ( dbo.Slice.SeriesSer = dbo.Series.SeriesSer ) and ( dbo.Series.SeriesUID = @UIDSerie )";
                using (SqlCommand command = new SqlCommand(cmdStr, conn))
                {
                    command.Parameters.AddWithValue("@UIDSerie", UIDSerie);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        try
                        {
                            couchVertPositionCT = (reader.GetDouble(0)) * (-10); // in mm
                            //MessageBox.Show("couchVertPositionCT: " + couchVertPositionCT.ToString());

                        }
                        catch (Exception)
                        {
                            reader.Close();
                        }

                    }
                }
                reader.Close();
                couchVertPositionCT = couchVertPositionCT - localContext.Collision.CouchVertPositionCTCorrection;

                return couchVertPositionCT;
            }
            catch (Exception e)
            {
                throw new Exception("couchVertPositionCT Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get the couch vertical position relative to the dicom origin (0,0,0) calculated from the upper face of the couch structure.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <returns>
        /// Couch vertical position in mm 
        /// </returns>
        public static double GetCouchVertPositionStructure(PlanSetup plan, LocalScriptContext localContext)
        {
            try
            {
                StructureSet structureSet = plan.StructureSet;
                string couchName = string.Empty;
                double couchVertPositionStructure = double.NaN;
                string machineId = plan.Beams.First().Id;
                //-- Get couch vertical position and couch name from structure --
                if (PlanCheckStructure.CheckCouchInsert(plan))
                {
                    foreach (var structure in structureSet.Structures)
                    {
                        if (structure.DicomType == "SUPPORT")
                        {
                            couchName = structure.Name;
                            couchVertPositionStructure = structure.CenterPoint.y;
                            break;
                        }
                    }
                }

                // -- The couch vertical position from structure is adjusted according to the name of the couch used.
                double correctionFactor; //mm       
                Machine m = localContext.Machines.Single(mach => mach.MachineId.Equals(machineId));
                CouchRegion c = m.Couch.CouchRegions.Find(item => item.CouchRegionName.Equals(couchName));
                correctionFactor = c.CouchVertPositionCorrection;
                couchVertPositionStructure = couchVertPositionStructure - correctionFactor; // Vertical Position of couch structure.
                //MessageBox.Show(couchVertPositionStructure.ToString());
                return couchVertPositionStructure;
            }
            catch (Exception e)
            {
                throw new Exception("GetCouchVertPositionStructure Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get the sector where the gantry moves to arrive at gantry angle position, 
        /// having account the extended range for Varian Auto Sequencie value. 
        /// Left(359.9-180.1) - Right (0-180)
        /// </summary>
        /// <param name="plan">Represents a treatment plan dataset.</param>
        /// <param name="b">Beam ID</param>
        /// <param name="conn">Open connection to a SQL Server database</param>
        /// <param name="gantry">gantry angle (degrees)</param>
        /// <returns>bool array with [0] = true -> left sector movement and [1] = true -> right sector movement</returns>
        public static bool[] GetRotationSector(PlanSetup plan, Beam b, SqlConnection conn, double gantry)
        {

            bool[] rotationSector = new bool[2]; // [0] = Gantry Left (359.9-180.1); [1] = Gantry Right (0-180);
            bool GR = false;
            bool GL = false;
            string planUID = plan.UID;
            string fieldId = b.Id;
            string extended = string.Empty;

            if (gantry >= 0 && gantry <= 180) // ARC or STATIC technique
            {
                GR = true;
            }
            else
            {
                GL = true;
            }

            // Get the extended range for Varian Auto Sequencie value.
            extended = PlanCheckBeam.GetExtendedFieldValue(b, plan, conn);

            //If the extended range for Varian Auto Sequence value is applied, the gantry rotation sector is interchanged.
            if (extended == "EN")
            {
                GR = !GR;
                GL = !GL;
            }

            rotationSector[0] = GL;
            rotationSector[1] = GR;

            return rotationSector;
        }


        /// <summary>
        /// Get the limit gantry angle where de collision with couch could be possible.
        /// </summary>
        /// <param name="b">Beam ID</param>
        /// <param name="imge"> Represents a CT dataset.</param>
        /// <param name="plan">Represents a treatment plan dataset.</param>
        /// <param name="couchVertPosition">couch vertical position relative to the dicom origin (0,0,0)</param>
        /// <param name="isocenter">isocenter position relativ to the dicom origin (0,0,0)</param>
        /// <param name="rotDirection">Gantry moves in Right sector (0-180 degrees) or Left sector (359.9-180.1 degrees)
        /// Values: Strings "Right" or "Left"</param>
        /// <param name="localContext"> Local Context</param>
        /// <returns>Limit gantry (degrees)</returns>
        public static double GetGantryLimitColl(Beam b, Image imge, PlanSetup plan, double couchVertPosition, VVector isocenter, string rotDirection, LocalScriptContext localContext)
        {
            try
            {
                double gantryLimit;
                double alpha;
                //double couchHigh = MachineVariables.couchCollisionParameters[0, 0]; // in mm (ExactCouch)
                Machine m = localContext.Machines.Single(mach => mach.MachineId.Equals("2100CD")); // in mm (ExactCouch) 
                double couchHigh = m.Couch.CouchRegions[1].CouchCollisionParameters[0].Item1;
                //double couchHalfWidth = MachineVariables.couchCollisionParameters[0, 1]; // in mm (ExactCouch)
                double couchHalfWidth = m.Couch.CouchRegions[1].CouchCollisionParameters[0].Item2; // in mm (ExactCouch)
                double TH = PlanCheckBeam.GetTh(couchVertPosition, isocenter);
                double isoX = isocenter.x;
                double isoY = isocenter.y;

                // MessageBox.Show("couchVertPosition: " + couchVertPosition + "\nisoY: " + isoY + "\nTH: " + TH + "\n" +"alpha: " + alpha);

                if (rotDirection == "Left")
                {
                    alpha = (Math.Atan((couchHalfWidth + isoX) / (TH + couchHigh))) * 180 / Math.PI;
                    gantryLimit = Math.Min(260, 215 + alpha);  // 360 - gantryLimit = 180 - 35 - alpha (35 is an aproximation to the angle between central axis and the gantry)
                }
                else if (rotDirection == "Right")
                {
                    alpha = (Math.Atan((couchHalfWidth - isoX) / (TH + couchHigh))) * 180 / Math.PI;
                    gantryLimit = Math.Max(100, 145 - alpha); // gantryLimit + 35 + alpha = 180
                }
                else
                {
                    throw new Exception("GetGantryLimitColl function Error: \n The gantry rotation direction is not defined correctly.");
                }

                return gantryLimit;
            }
            catch (Exception e)
            {
                throw new Exception("GetGantryLimitColl function Error: \n" + e.Message);
            }


        }


        /// <summary>
        /// Check if gantry collides with patient or couch.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <param name="imge">
        /// Represents a CT dataset.
        /// </param>
        /// <param name="couchVertPosition">
        /// Couch vertical position in mm
        /// </param>
        /// <returns>
        /// Tuple <string,bool,bool> where Item1 -> Collision message, Item2 -> true if collision with couch, Item3 -> true if collision with patient.
        /// </returns>
        public static Tuple<string, bool, bool> CheckCollisions(PlanSetup plan, LocalScriptContext localContext, Image imge, double couchVertPosition)
        {
            // Tuple<Warning message Text, collision couch, collision patient>            
            bool CouchRotOutOfRange = false;
            double[] patientCollision = new double[3]; // [0] = GantryLimitPatient; [1] = gantryMargin (degrees); [2] = Collision Level ((0) No collision, (1) warning, (2) collision)
            double[] couchCollision = new double[4]; // [0] = GantryLimitCouch; [1] = couchMarginCollision (mm); [2] = Collision Level ((0) No collision, (1) warning, (2) collision); [3] -> Gantry Angle (degrees) 
            List<Tuple<string, int, double[]>> patientCollisionList = new List<Tuple<string, int, double[]>>(); // string -> BeamID ; int -> 0 = Start gantry Angle / 1 = End Gantry Angle
            // double [] -> patientCollision
            List<Tuple<string, int, double[]>> couchCollisionList = new List<Tuple<string, int, double[]>>();
            bool[] gantryRotSector = new bool[2]; // [0] = Gantry Left (359.9-180.1); [1] = Gantry Right (0-180); 
            string msgNotCalc = string.Empty;
            int count = 0;
            double startGantryAngle = double.NaN;
            double endGantryAngle = double.NaN;
            int nControlPoints;
            int nSubField;
            SqlConnection conn = localContext.MySQLConnection.Conn;

            foreach (Beam b in plan.Beams) // collision for each field.
            {

                if (localContext.Collision.MaxCouchRotCalc >= (b.ControlPoints[0].PatientSupportAngle) || 360 - localContext.Collision.MaxCouchRotCalc <= (b.ControlPoints[0].PatientSupportAngle))
                {
                    VVector isocenter = PlanCheckBeam.GetIsocenter(b, imge, plan); // Get isocenter corrected by image(CT) and patient(plan) orientation.

                    if (b.Technique.Id.ToString() == "ARC")
                    {
                        nSubField = 2;
                        startGantryAngle = b.ControlPoints[0].GantryAngle;  // Start Gantry Angle (degrees)
                        nControlPoints = b.ControlPoints.Count();
                        endGantryAngle = b.ControlPoints[nControlPoints - 1].GantryAngle; // End Gantry Angle (degrees)
                    }
                    else
                    {
                        nSubField = 1;
                        startGantryAngle = b.ControlPoints[0].GantryAngle;  // Start Gantry Angle (degrees)
                    }

                    for (int i = 0; i < nSubField; i++)
                    {
                        if (i == 0) // Start Gantry Angle.
                        {
                            gantryRotSector = PlanCheckCollisions.GetRotationSector(plan, b, conn, startGantryAngle); // bool -> [0] = Gantry Left (359.9-180.1); [1] = Gantry Right (0-180);
                            patientCollision = PlanCheckCollisions.CheckCollisionPatient(b, imge, plan, couchVertPosition, startGantryAngle, gantryRotSector, isocenter);
                            couchCollision = PlanCheckCollisions.CheckCollisionCouch(b, imge, plan, localContext, couchVertPosition, startGantryAngle, gantryRotSector, isocenter);
                        }
                        else if (i == 1) // End Gantry Angle.
                        {
                            gantryRotSector = PlanCheckCollisions.GetRotationSector(plan, b, conn, endGantryAngle); // bool -> [0] = Gantry Left (359.9-180.1); [1] = Gantry Right (0-180);
                            patientCollision = PlanCheckCollisions.CheckCollisionPatient(b, imge, plan, couchVertPosition, endGantryAngle, gantryRotSector, isocenter);
                            couchCollision = PlanCheckCollisions.CheckCollisionCouch(b, imge, plan, localContext, couchVertPosition, endGantryAngle, gantryRotSector, isocenter);
                        }

                        // List with patient and couch collision results.
                        patientCollisionList.Add(Tuple.Create(b.Id.ToString(), i, patientCollision));
                        couchCollisionList.Add(Tuple.Create(b.Id.ToString(), i, couchCollision));

                        //MessageBox.Show("listcouch: " + "beam ID: " + couchCollisionList[count].Item1 + " || Start or end: " + couchCollisionList[count].Item2 + " || Gantry Limit:  " + couchCollisionList[count].Item3[0].ToString() + " || Margin:  " + couchCollisionList[count].Item3[1].ToString() + " || Collision Level:  " + couchCollisionList[count].Item3[2].ToString() + "\n");
                        //MessageBox.Show("listPatient: " + "beam ID: " + patientCollisionList[count].Item1 + " || Start or end: " + patientCollisionList[count].Item2 + " || Gantry Limit:  " + patientCollisionList[count].Item3[0].ToString() + " || Margin:  " + patientCollisionList[count].Item3[1].ToString() + " || Collision Level:  " + patientCollisionList[count].Item3[2].ToString() + "\n");

                        count = count + 1;  // How many fields will be analized (have account Start and End fields)

                    }
                }
                else
                {
                    if (CouchRotOutOfRange == false)
                    {
                        msgNotCalc += ("Collision not evaluated (Couch Rotation > " + localContext.Collision.MaxCouchRotCalc + " degrees): \n");
                        CouchRotOutOfRange = true;
                    }
                    msgNotCalc += ("          - " + b.Id.ToString() + "\n");
                }
            }

            // ---- Print messages ----            
            Tuple<bool, bool, string> printColl = PlanCheckCollisions.PrintCollision(patientCollisionList, couchCollisionList, count, localContext);

            // Tuple<Warning message Text, collision couch, collision patient>
            Tuple<string, bool, bool> collisions = new Tuple<string, bool, bool>(printColl.Item3 + msgNotCalc, printColl.Item2, printColl.Item1);

            return collisions;
        }


        /// <summary>
        /// Check gantry-patient collision.
        /// </summary>
        /// <param name="b">Beam ID</param>
        /// <param name="imge"> Represents a CT dataset.</param>
        /// <param name="plan">Represents a treatment plan dataset.</param>
        /// <param name="conn">Open connection to a SQL Server database</param>
        /// <param name="couchVertPosition">couch vertical position relative to the dicom origin (0,0,0)</param>
        /// <param name="gantryAngle">gantry angle (degrees)</param>
        /// <param name="gantryRotSector">bool array where 
        ///                                 [0] = true - left sector movement 
        ///                                 [1] = true - right sector movement 
        /// </param>
        /// <param name="isocenter">isocenter position relativ to the dicom origin (0,0,0)</param>
        /// <returns>   double array where 
        ///             [0] -> Gantry limit patient collision value (degrees)
        ///             [1] -> Gantry margin for collision (degrees)
        ///             [2] -> Collision level: (0) -> No collision ; (1) -> Warning ; (2) -> Collision             
        /// </returns>

        public static double[] CheckCollisionPatient(Beam b, Image imge, PlanSetup plan, double couchVertPosition, double gantryAngle, bool[] gantryRotSector, VVector isocenter)
        {
            double[] collisionPatient = new double[3]; // [0] = GantryLimitPatient; [1] = gantryMargin (degrees); [2] -> Collision level
            double limitGantry = double.NaN;
            double marginGantry = double.NaN;
            bool GL = gantryRotSector[0]; // Left Sector
            bool GR = gantryRotSector[1]; // Right Sector

            try
            {
                double isoX = isocenter.x;
                double isoY = isocenter.y;
                double couchRot = b.ControlPoints[0].PatientSupportAngle;

                double a = 260 + Math.Abs(isoX);
                double aRot = a / Math.Cos(couchRot * Math.PI / 180) + 40 * Math.Abs(Math.Tan(couchRot * Math.PI / 180));
                double aDif = aRot - a;

                if (isoX < -70)
                {
                    if ((gantryAngle > 35 && GR == true))
                    {
                        limitGantry = (565 + (PlanCheckBeam.GetTh(couchVertPosition, isocenter)) * 0.65 + (isoX - aDif) * 1.1) / 10;
                        marginGantry = limitGantry - gantryAngle;
                    }
                }
                else if (isoX > 70)
                {
                    if (gantryAngle < 325 && GL == true)
                    {
                        limitGantry = (3600 - (565 + (PlanCheckBeam.GetTh(couchVertPosition, isocenter)) * 0.65 - (isoX + aDif) * 1.1)) / 10;
                        marginGantry = gantryAngle - limitGantry;
                    }
                }

                // MessageBox.Show("TH: " + Math.Round((couchVertPosition - isoY), 2).ToString() + " isoX: " + Math.Round(isoX, 2).ToString() + "\n" + "gantryAngle: " + gantryAngle.ToString() + " - limitGantry/10: " + Math.Round((limitGantry / 10), 2).ToString() + " = marginGantryGantry: " + Math.Round(marginGantry, 2).ToString());

                double gantryWarning = 2; // margin for gantry angle (deg) for warning purposes


                collisionPatient[0] = double.NaN;
                collisionPatient[1] = double.NaN;
                collisionPatient[2] = 0; // No collision

                if (marginGantry <= 0)
                {
                    collisionPatient[0] = limitGantry;
                    collisionPatient[1] = marginGantry;
                    collisionPatient[2] = 2; // Collision.

                }
                else if ((GR == true && gantryAngle > (limitGantry - gantryWarning) && gantryAngle < limitGantry) ||
                    (GL == true && gantryAngle < (limitGantry + gantryWarning) && gantryAngle > limitGantry))
                {
                    collisionPatient[0] = limitGantry;
                    collisionPatient[1] = double.NaN;
                    collisionPatient[2] = 1; // No collision but Warning
                }

                return collisionPatient;
            }
            catch (Exception e)
            {
                throw new Exception("PatientMarginGantryCalc function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Check gantry-couch collision
        /// </summary>
        /// <param name="b">Beam ID</param>
        /// <param name="imge"> Represents a CT dataset.</param>
        /// <param name="plan">Represents a treatment plan dataset.</param>
        /// <param name="conn">Open connection to a SQL Server database</param>
        /// <param name="couchVertPosition">couch vertical position relative to the dicom origin (0,0,0)</param>
        /// <param name="gantryAngle">gantry angle (degrees)</param>
        /// <param name="gantryRotSector">bool array where 
        ///                                 [0] = true - left sector movement 
        ///                                 [1] = true - right sector movement 
        /// </param>
        /// <param name="isocenter">isocenter position relativ to the dicom origin (0,0,0)</param>
        /// <returns>   double array where 
        ///             [0] -> Gantry limit couch collision value (degrees)
        ///             [1] -> Margin for couch collision (mm)
        ///             [2] -> Collision level: (0) -> No collision ; (1) -> Warning ; (2) -> Collision
        ///             [3] -> Gantry Angle (degrees)
        /// </returns>
        public static double[] CheckCollisionCouch(Beam b, Image imge, PlanSetup plan, LocalScriptContext localContext, double couchVertPosition, double gantryAngle, bool[] gantryRotSector, VVector isocenter)
        {
            double[] couchCollision = new double[4];

            // Get the gantry limit angle for collision with couch.
            double gantryLimitRight = PlanCheckCollisions.GetGantryLimitColl(b, imge, plan, couchVertPosition, isocenter, "Right", localContext);
            double gantryLimitLeft = PlanCheckCollisions.GetGantryLimitColl(b, imge, plan, couchVertPosition, isocenter, "Left", localContext);

            double gantryLimit = double.NaN;
            string machineName = b.TreatmentUnit.Id;
            double couchRot = b.ControlPoints[0].PatientSupportAngle;

            bool GL = gantryRotSector[0];
            bool GR = gantryRotSector[1];

            //double margin = -1000;
            double margin = double.NaN;
            double margin1, margin2;
            double d1, d2;
            double a1, b1, r1, a2, b2, r2;

            try
            {
                // Calculation of margin for machines with Exact Couch.
                Machine m = localContext.Machines.Single(mach => mach.MachineId.Equals(machineName));                
               
                //if (Array.IndexOf(MachineVariables.machinesExactCouch, machineName) != -1)
                if (m.Couch.CouchName == "ExactCouch")
                {
                    a1 = m.Couch.CouchRegions[0].CouchCollisionParameters[0].Item1; //in mm
                    b1 = m.Couch.CouchRegions[0].CouchCollisionParameters[0].Item2; //in mm         El valor b1 en VB es 240.
                    r1 = m.Couch.CouchRegions[0].CouchCollisionParameters[0].Item3; //in mm
                    d1 = double.NaN;

                    if (GL == true)
                    {
                        b1 = b1 - isocenter.x * (-1);
                        gantryLimit = gantryLimitLeft;
                    }
                    else if (GR == true)
                    {
                        b1 = b1 - isocenter.x;
                        gantryLimit = gantryLimitRight;
                    }

                    b1 = b1 / Math.Cos(couchRot * Math.PI / 180) + 40 * Math.Abs(Math.Tan(couchRot * Math.PI / 180));

                    d1 = Math.Sqrt(Math.Pow((a1 + PlanCheckBeam.GetTh(couchVertPosition, isocenter)), 2) +
                                Math.Pow(b1, 2));

                    //MessageBox.Show("d1:" + d1);

                    margin = r1 - d1;
                }

                // Calculation of margin for machines with Exact IGRT Couch.
                // else if (Array.IndexOf(MachineVariables.machinesExactIGRTCouch, machineName) != -1)
                else if (m.Couch.CouchName == "IGRTCouch")
                {
                    a1 = m.Couch.CouchRegions[0].CouchCollisionParameters[0].Item1; //in mm
                    b1 = m.Couch.CouchRegions[0].CouchCollisionParameters[0].Item2; //in mm         El valor b1 en VB es 240.
                    r1 = m.Couch.CouchRegions[0].CouchCollisionParameters[0].Item3; //in mm

                    a2 = m.Couch.CouchRegions[0].CouchCollisionParameters[1].Item1; //in mm
                    b2 = m.Couch.CouchRegions[0].CouchCollisionParameters[1].Item2; //in mm         El valor b1 en VB es 240.
                    r2 = m.Couch.CouchRegions[0].CouchCollisionParameters[1].Item3; //in mm


                    d1 = d2 = double.NaN; // in mm

                    b1 = b1 / Math.Cos(couchRot * Math.PI / 180) + 40 * Math.Abs(Math.Tan(couchRot * Math.PI / 180));
                    b2 = b2 / Math.Cos(couchRot * Math.PI / 180) + 40 * Math.Abs(Math.Tan(couchRot * Math.PI / 180));

                    if (GL == true)
                    {
                        b1 = b1 - isocenter.x * (-1);
                        b2 = b2 - isocenter.x * (-1);
                        gantryLimit = gantryLimitLeft;
                    }
                    else if (GR == true)
                    {
                        b1 = b1 - isocenter.x;
                        b2 = b2 - isocenter.x;
                        gantryLimit = gantryLimitRight;
                    }

                    d1 = Math.Sqrt(Math.Pow((a1 + PlanCheckBeam.GetTh(couchVertPosition, isocenter)), 2) +
                            Math.Pow(b1, 2));
                    d2 = Math.Sqrt(Math.Pow((a2 + PlanCheckBeam.GetTh(couchVertPosition, isocenter)), 2) +
                            Math.Pow(b2, 2));

                    //MessageBox.Show("d1: " + d1 + " -- d2: "+ d2);
                    margin1 = r1 - d1;
                    margin2 = r2 - d2;

                    margin = Math.Min(margin1, margin2);
                }
                else
                {
                    throw new Exception("Machine without Couch type assigned");
                }


                // Collision level 

                if ((GR == true && gantryAngle < (gantryLimitRight - localContext.Collision.CollisionSafetyMarginGantryAngle)) || (GL == true && gantryAngle > (gantryLimitLeft + localContext.Collision.CollisionSafetyMarginGantryAngle)))
                {
                    couchCollision[0] = double.NaN;
                    couchCollision[1] = double.NaN;
                    couchCollision[2] = 0;  //No collision
                }
                else if (((GR == true && gantryAngle >= (gantryLimitRight - localContext.Collision.CollisionSafetyMarginGantryAngle) && gantryAngle < gantryLimitRight) ||
                        (GL == true && gantryAngle <= (gantryLimitLeft + localContext.Collision.CollisionSafetyMarginGantryAngle) && gantryAngle > gantryLimitLeft)) && margin < 0)
                {
                    couchCollision[0] = gantryLimit;
                    couchCollision[1] = margin;
                    couchCollision[2] = 1; // No collision but Warning. Warning with margin <0 when gantry close to gantryLimit. "Risk of collision due to gantry angle"
                }
                else if (((GR == true && gantryAngle > gantryLimitRight) || (GL == true && gantryAngle < gantryLimitLeft)) && margin <= 0)
                {
                    couchCollision[0] = gantryLimit;
                    couchCollision[1] = margin;
                    couchCollision[2] = 2; //Collision
                }
                else if (((GR == true && gantryAngle > gantryLimitRight) || (GL == true && gantryAngle < gantryLimitLeft)) && (margin > 0 && margin <= localContext.Collision.CollisionSafetyMarginDistance))
                {
                    couchCollision[0] = gantryLimit;
                    couchCollision[1] = margin;
                    couchCollision[2] = 1; // No collision but Warning. Warning with margin >0 when gantry over gantryLimit and margin close to zero. "Risk of collision due to reduced margin"
                }

                couchCollision[3] = gantryAngle;

                return couchCollision;
            }
            catch (Exception e)
            {
                throw new Exception("CheckCollisionCouch function Error: \n" + e.Message);
            }
        }


        /// <summary>
        /// Print the gantry collision with patient and couch results 
        /// </summary>
        /// <param name="patientCollisionList">
        /// List with a patient collision Tuple where
        ///         patientCollisionList ---> [Item1] -> BeamID ; [Item2]-> 0 = Start gantry Angle / 1 = End Gantry Angle ; [Item3]double [] -> patientCollision
        ///         patientCollision --->  [0] = GantryLimitPatient; [1] = gantryMargin (degrees); [2] = Collision Level ((0) No collision, (1) warning, (2) collision)
        /// </param>
        /// <param name="couchCollisionList">
        /// List with a couch collision Tuple where
        ///         couchCollisionList ---> [Item1]string -> BeamID ; [Item2]int -> 0 = Start gantry Angle / 1 = End Gantry Angle ; [Item3]double [] -> couchCollision
        ///         couchCollision ---> [0] = GantryLimitCouch; [1] = couchMarginCollision (mm); [2] = Collision Level ((0) No collision, (1) warning, (2) collision); [3] -> Gantry Angle (degrees)
        /// </param>
        /// <param name="count">
        /// Number of fields with collision
        /// </param>
        /// <param name="localContext">
        /// Local Context
        /// </param>
        /// <returns>
        /// Tuple with collision message and bool indications of collision.
        ///     Tuple ---> [Item1] -> True if not collision patient / [Item2] -> True if not collision couch / [Item3] -> Collision message.
        /// </returns>
        public static Tuple<bool, bool, string> PrintCollision(List<Tuple<string, int, double[]>> patientCollisionList, List<Tuple<string, int, double[]>> couchCollisionList, int count, LocalScriptContext localContext)
        {

            string msg = string.Empty;
            bool collisionPatient = false;
            bool collisionCouch = false;

            // patientCollisionList ---> [Item1]string -> BeamID ; [Item2]int -> 0 = Start gantry Angle / 1 = End Gantry Angle ; [Item3]double [] -> patientCollision
            // patientCollision --->  [0] = GantryLimitPatient; [1] = gantryMargin (degrees); [2] = Collision Level ((0) No collision, (1) warning, (2) collision)
            string starEndGantry = "";
            
            if (patientCollisionList.Exists(x => x.Item3[2] == 2)) // Check if collision with patient exists.
            {
                msg += "          -   COLLISION with PATIENT:\n";
                collisionPatient = false;

                for (int i = 0; i < count; i++) // Print every field with patient collision.
                {
                    if (patientCollisionList[i].Item3[2] == 2)
                    {
                        if (couchCollisionList[i].Item2 == 0) { starEndGantry = "Start G"; } else { starEndGantry = "End G"; } // Get the label "Start G" or "End G"

                        msg += "\t" + (patientCollisionList[i].Item1 + " (" + starEndGantry + ")" + "\tLimit: " + Math.Round(patientCollisionList[i].Item3[0], 1) + " deg\n"); // patientCollision[0] = Collision grantry value (degrees)
                    }
                }

            }
            else
            {
                collisionPatient = true;
            }

            if (patientCollisionList.Exists(x => x.Item3[2] == 1)) // Check if warning collision with patient exists.
            {
                msg += ("          -   Gantry close to PATIENT: \n");

                for (int i = 0; i < count; i++) // Print every field with patient warning collision
                {
                    if (patientCollisionList[i].Item3[2] == 1)
                    {
                        if (couchCollisionList[i].Item2 == 0) { starEndGantry = "Start G"; } else { starEndGantry = "End G"; } // Get the label "Start G" or "End G"

                        msg += "\t" + (patientCollisionList[i].Item1 + " (" + starEndGantry + ")" + "\tLimit: " + Math.Round(patientCollisionList[i].Item3[0], 1) + " deg\n"); // patientCollision[0] = Collision grantry value (degrees)
                    }
                }

            }

            // couchCollisionList ---> [Item1]string -> BeamID ; [Item2]int -> 0 = Start gantry Angle / 1 = End Gantry Angle ; [Item3]double [] -> couchCollision
            // couchCollision ---> [0] = GantryLimitCouch; [1] = couchMarginCollision (mm); [2] = Collision Level ((0) No collision, (1) warning, (2) collision); [3] -> Gantry Angle (degrees)
            // Couch Collision
            if (couchCollisionList.Exists(x => x.Item3[2] == 2)) // Check if collision with couch exists.
            {
                msg += "          -   COLLISION with couch: \n";
                collisionCouch = false;

                for (int i = 0; i < count; i++) // Print every field with couch collision.
                {

                    if (couchCollisionList[i].Item3[2] == 2)
                    {
                        if (couchCollisionList[i].Item2 == 0) { starEndGantry = "Start G"; } else { starEndGantry = "End G"; } // Get the label "Start G" or "End G"

                        if (Math.Abs(Math.Round(couchCollisionList[i].Item3[0], 1) - Math.Round(couchCollisionList[i].Item3[3], 1)) < localContext.Collision.CollisionSafetyMarginGantryAngle)
                        {
                            msg += "\t" + (couchCollisionList[i].Item1 + " (" + starEndGantry + ")" + "\tMargin: " + Math.Round(couchCollisionList[i].Item3[1] / 10, 1).ToString() + " cm\t(Limit: " + Math.Round(couchCollisionList[i].Item3[0], 1) + " deg)\n"); // couchCollision[0] = Collision grantry value (degrees)
                        }
                        else
                        {
                            msg += "\t" + (couchCollisionList[i].Item1 + " (" + starEndGantry + ")" + "\tMargin: " + Math.Round(couchCollisionList[i].Item3[1] / 10, 1).ToString() + " cm\n"); // couchCollision[0] = Collision grantry value (degrees)
                        }

                    }
                }

            }
            else
            {
                collisionCouch = true;
            }

            bool riskCollGantry = false;
            bool riskCollMargin = false;
            if (couchCollisionList.Exists(x => x.Item3[2] == 1)) // Check if warning collision with couch exists.
            {
                for (int i = 0; i < count; i++) // Print every field with "Risk of collision due to gantry angle"
                {
                    if (couchCollisionList[i].Item3[2] == 1 && couchCollisionList[i].Item3[1] < 0)
                    {
                        if (!riskCollGantry)
                        {
                            msg += ("          -   Gantry near couch: \n");
                            riskCollGantry = true;
                        }

                        if (couchCollisionList[i].Item2 == 0) { starEndGantry = "Start G"; } else { starEndGantry = "End G"; } // Get the label "Start G" or "End G"

                        msg += "\t" + (couchCollisionList[i].Item1 + " (" + starEndGantry + ")" + "\tLimit: " + Math.Round(couchCollisionList[i].Item3[0], 1) + " deg\n"); // couchCollision[0] = Collision grantry value (degrees)
                    }
                }



                for (int i = 0; i < count; i++) // Print every field with "Risk of collision due to reduced margin"
                {
                    if (couchCollisionList[i].Item3[2] == 1 && couchCollisionList[i].Item3[1] >= 0 && couchCollisionList[i].Item3[1] <= localContext.Collision.CollisionSafetyMarginDistance)
                    {

                        if (!riskCollMargin)
                        {
                            msg += ("          -   Reduced margin to couch: \n");
                            riskCollMargin = true;
                        }
                        if (couchCollisionList[i].Item2 == 0) { starEndGantry = "Start G"; } else { starEndGantry = "End G"; } // Get the label "Start G" or "End G"

                        if (Math.Abs(Math.Round(couchCollisionList[i].Item3[0], 1) - Math.Round(couchCollisionList[i].Item3[3], 1)) < localContext.Collision.CollisionSafetyMarginGantryAngle) // (gantryLimit - gantryAngle) < 5 degrees
                        {
                            msg += "\t" + (couchCollisionList[i].Item1 + " (" + starEndGantry + ")" + "\tMargin: " + Math.Round(couchCollisionList[i].Item3[1] / 10, 1).ToString() + " cm\t(Limit: " + Math.Round(couchCollisionList[i].Item3[0], 1) + " deg)\n"); // couchCollision[0] = Collision grantry value (degrees)
                        }
                        else
                        {
                            msg += "\t" + (couchCollisionList[i].Item1 + " (" + starEndGantry + ")" + "\tMargin: " + Math.Round(couchCollisionList[i].Item3[1] / 10, 1).ToString() + " cm\n"); // couchCollision[0] = Collision grantry value (degrees)
                        }
                    }
                }

            }


            if (msg == string.Empty)
            {
                collisionCouch = true;
                collisionPatient = true;
            }
            else
            {
                msg = "Collisions: \n" + msg;
            }

            //MessageBox.Show("collisionCouch: " + collisionCouch);
            //MessageBox.Show("collisionPatient: " + collisionPatient);
            //MessageBox.Show(msg);
            Tuple<bool, bool, string> collTuple = new Tuple<bool, bool, string>(collisionPatient, collisionCouch, msg);


            //MessageBox.Show("Collision Check --->  " + msg);
            return collTuple;


        }

    }
    // ----------------------------------------------------------------------
    // ---------------------   End class PlanCheckCollisions ----------------
    // ----------------------------------------------------------------------
    #endregion

    #region VERIFICATION FUNCTIONS

    /// <summary>
    /// Set of functions to verify plan.
    /// </summary>
    public class PlanCheckVerification
    {
        /// <summary>
        /// Check initial parameters to run the Script.
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
        /// Out message with errors.
        /// </param>
        /// <returns>
        /// True if some error exists
        /// </returns>
        public static bool CheckInitialParameters(ScriptContext context, PlanSetup plan, Image imge, out string msg)
        {
            
            bool error = false;
            msg = string.Empty;
            try
            {
                // If there's no selected plan with calculated dose throw an exception
                if (plan == null && context.PlanSumsInScope.Count()>0)
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
                        PlanCheckPlanPrescription.GetNumberOfFracctions(plan);
                        PlanCheckPlanPrescription.GetDosePerFracction(plan);
                        PlanCheckPlanPrescription.GetTotalDose(plan);
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

    #region SQL FUNCTIONS

    /// <summary>
    /// Set of SQL functions
    /// </summary>
    public class PlanCheckSQL
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
                "Data Source=" + dataSource + ";" +
                "Initial Catalog=" + initialCatalog + ";" +
                "User id=" + userID + ";" +
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

    #region AUXILIAR FUNCTIONS

    /// <summary>
    /// Set of auxiliar functions
    /// </summary>
    public class PlanCheckAux
    {
        /// <summary>
        /// Execute an external application
        /// </summary>
        /// <param name="App">Application name if it is in system path, if not the full path</param>
        /// <param name="Args">Arguments for application</param>
        /// <returns>true if the execution is successful</returns>
        public static bool ExecuteProgram(String App, String Args)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = App;
                process.StartInfo.Arguments = Args;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
                process.Start();
                process.WaitForExit();
                //MessageBox.Show(process.ExitCode.ToString());
                return (process.ExitCode == 0) ? true : false;
            }
            catch (Exception e)
            {
                throw new Exception("ExecuteProgram function Error: \n" + e.Message);
            }

        }


    }


    //  -----------------------------   End Auxilar Functions -----------------------------------------
    #endregion

    #region STRING EXTENSIONS

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
    /// Set of printing functions
    /// </summary>
    public class PlanCheckPrinting
    {
        //"\u2714" is the "ok" symbol in unicode
        //"\u2718" is the "x" symbol in unicode
        //"\u0020" is the "space" symbol in unicode
        //"\u2022" is the "bullet" symbol in unicode

        /// <summary>
        /// Print the plan verifation message
        /// </summary>
        /// <param name="verificationTable">
        /// Data table with the verificacion messages.
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
        /// Print the plan Warning message
        /// </summary>
        /// <param name="warningTable">
        /// Data table with the Warning messages.
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
        /// Print the plan information message
        /// </summary>
        /// <param name="informationTable">
        /// Data table with the Information messages.
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