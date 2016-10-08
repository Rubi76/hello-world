using System;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Collections;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;
using System.Globalization;

namespace VMS.TPS
{

    // Esto es una modificación local
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

    #region --- SCRIPT ---
    /// <summary>
    ///  Scripting Plan Check
    /// </summary>
    public class Script
    {
        /// <summary>
        /// Script constructor without paramameters.
        /// </summary>
        public Script()
        {
        }
        
        public static PlanningItem myPlan;
        public static IEnumerable<Structure> myStructures;
        public static Fractionation myPrescription;
        

        /// <summary>
        /// Constraints analyzed messages - Table definition
        /// </summary>
        public static DataTable constraintsAnalyzed = new DataTable();
        /// <summary>
        /// Constraints Errors messages - Table definition
        /// </summary>
        public static DataTable constraintsErrors = new DataTable();

        /// <summary>
        /// Script (Main Program) 
        /// </summary>
        /// <param name="context">
        /// The TPS context.
        /// </param>
        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {

            // Initial checks  
            string initialMsg = string.Empty;
            LocalScriptContext localContext = null;

            // Local Context generation
            try
            {
                localContext = LocalScriptContextGeneration.localScriptContextGeneration();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message); //Exit Script
                //MessageBox.Show(e.Message);
            }
            
            // Inital verifications
            if (Verifications.CheckInitialParameters(context, out initialMsg))
            {
                throw new Exception(initialMsg); //Exit Script
            }

            try {
                // Initialization of variables depending on whether the plan is individual or sum.  
                if (context.PlanSetup != null)
                {
                    myPlan = context.ExternalPlanSetup;
                    myStructures = ((ExternalPlanSetup)myPlan).StructureSet.Structures;
                    myPrescription = ((ExternalPlanSetup)myPlan).UniqueFractionation;
                }
                else // Atenció només accepta el "primer" PlanSum
                {
                    myPlan = context.PlanSumsInScope.FirstOrDefault();
                    myStructures = context.PlanSumsInScope.FirstOrDefault().StructureSet.Structures;
                    myPrescription = context.PlanSumsInScope.FirstOrDefault().PlanSetups.FirstOrDefault().UniqueFractionation;
                }

                // Constraints Analyzed Table      
                // nombre OR, objetivo_rtPrescription, objetivo_conseguido, check_OK_False
                Script.constraintsAnalyzed.Columns.Add("ORName", typeof(string));
                Script.constraintsAnalyzed.Columns.Add("RTPrescriptionObjectiveLabel", typeof(string));
                Script.constraintsAnalyzed.Columns.Add("RTPrescriptionObjectivePrint", typeof(string));
                Script.constraintsAnalyzed.Columns.Add("ObjectiveAchievedValue", typeof(string));
                Script.constraintsAnalyzed.Columns.Add("ObjectiveAchievedBool", typeof(bool));

                // Constraints Errors Table      
                Script.constraintsErrors.Columns.Add("ConstraintErrors", typeof(string));
            }
            catch (Exception e)
            {
                throw new Exception(e.Message); //Exit Script
            }


            //tuple<ItemValue,ItemType,AnatomyRole,AnatomyName,IntemValueUnit,PrescriptionName,Param1Value,Param1ValueUnit>
            List<Tuple<string, string, int, string, string, string, string, Tuple<string>>> rtPrescription = new List<Tuple<string, string, int, string, string, string, string, Tuple<string>>>();
            try
            {               
                rtPrescription = RTPrescription.GetRTPrescription(context, myPlan, localContext);               
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }         
            
            try
            {

                if (!rtPrescription.Any(t => t.Item3 == 3))
                {
                    throw new Exception("There are no constraints");
                }

                // Check constraints
                Constraints.CheckConstraints(rtPrescription, myStructures, myPlan, myPrescription);

                // Print results                
                CheckPrinting.PrintConstraintsAnalyzed(constraintsAnalyzed, myPlan);
                CheckPrinting.PrintConstraintWarnings(constraintsErrors);

                // --------- Close SQL connection ---------------
               
                PlanCheckSQL.SQLCloseConnection(localContext.MySQLConnection.Conn);

            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            

            //////Structure str = null;
            //////foreach (Structure s in myStructures)
            //////{
            //////    //MessageBox.Show(s.Id);
            //////    if (s.Id == rtPrescription[10].Item4 && rtPrescription[10].Item3 == 3)
            //////    {
            //////        str = s;
            //////        MessageBox.Show(str.Id);
            //////        break;
            //////    }
            //////}

            //////// Get mean dose
            //////DoseValue meanDoseRelative = CheckDoseVolume.GetMeanDose((PlanSetup)myPlan, "Relative", str);
            //////DoseValue meanDoseAbsolute = CheckDoseVolume.GetMeanDose((PlanSetup)myPlan, "Absolute", str);
            //////MessageBox.Show("meanDoseRelative: " + meanDoseRelative.Dose + " %\n" + "meanDoseAbsolute: " + meanDoseAbsolute.Dose + " Gy");

            //////// Get Max dose
            //////DoseValue maxDoseRelative = CheckDoseVolume.GetMaxDose((PlanSetup)myPlan, "Relative", str);
            //////DoseValue maxDoseAbsolute = CheckDoseVolume.GetMaxDose((PlanSetup)myPlan, "Absolute", str);
            //////MessageBox.Show("maxDoseRelative: " + maxDoseRelative.Dose + " %\n" + "maxDoseAbsolute: " + maxDoseAbsolute.Dose + " Gy");

            //////// (14Gy -> x % Volume)
            //////double V14 = CheckDoseVolume.GetPercentVolume((PlanSetup)myPlan, str, 14, "Absolute");
            ////////GetPercentVolume(PlanSetup plan, Structure str, double value, string presentation)
            //////MessageBox.Show("V14: " + V14 + " %");

            //////// (95% total dose -> x % volume)
            //////double v95percent = CheckDoseVolume.GetPercentVolume((PlanSetup)myPlan, str, 95,"Relative");
            //////MessageBox.Show("v95%: " + v95percent +" %");

            //////// (25% Volume -> x Gy)
            //////DoseValue pv25Percent = CheckDoseVolume.GetAbsoluteDose((PlanSetup)myPlan, str, 25, "Relative");
            //////MessageBox.Show("pv25%: " + pv25Percent.Dose + " Gy");

            //////// (10cm3 -> x Gy)
            //////DoseValue pv10cm3 = CheckDoseVolume.GetAbsoluteDose((PlanSetup)myPlan, str, 10, "AbsoluteCm3");
            //////MessageBox.Show("pv10cm3: " + pv10cm3.Dose + " Gy");

            //double myV95 = ((PlanSetup)myPlan).GetVolumeAtDose(str, new DoseValue(0.95 * 15, DoseValue.DoseUnit.Gy), VolumePresentation.Relative); // Porcentaje de volumen que recive el 95% de 15Gy
            //MessageBox.Show(myV95.ToString());

            //DVHData myDVHabs = myPlan.GetDVHCumulativeData(myTarget, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.001);
            //DVHData myDVHrel = myPlan.GetDVHCumulativeData(myTarget, DoseValuePresentation.Relative, VolumePresentation.Relative, 0.001);

            //double myD2abs = myPlan.GetDoseAtVolume(myTarget, 2.0, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
            //double myD2rel = myPlan.GetDoseAtVolume(myTarget, 2.0, VolumePresentation.Relative, DoseValuePresentation.Relative).Dose / doseScaling;
            //double myD98abs = myPlan.GetDoseAtVolume(myTarget, 98.0, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
            //double myD98rel = myPlan.GetDoseAtVolume(myTarget, 98.0, VolumePresentation.Relative, DoseValuePresentation.Relative).Dose / doseScaling;


        }
    }
    #endregion

    #region RT PRESCRIPTION
    public class RTPrescription
    {
        /// <summary>
        /// Get a full RT Prescription for a specific plan or plan sum  (Physician's intent).
        /// - Prescription (VolumeId, dose per fraction, total dose)
        /// - Volume dose coverage (maximum dose, minimum dose, minimum dvh dose, maximum dvh dose)
        /// - Constraints (maximum mean dose, maximum dose, OAR FREE TEXT)
        /// </summary>
        /// <param name="context">
        /// The TPS context.
        /// </param>
        /// <param name="plan">
        /// Plan Setup or Plan Sum
        /// </param>
        /// <param name="localContext">
        /// Represent a local context
        /// </param>
        /// <returns>
        /// List of Tuples where a full RT Prescription is stored.
        /// Tuple -> ItemValue,ItemType,AnatomyRole,AnatomyName,IntemValueUnit,PrescriptionName,Param1Value,Param1ValueUnit
        /// For more information of each item in tuple see "data base information
        /// </returns>
        public static List<Tuple<string, string, int, string, string, string, string, Tuple<string>>> GetRTPrescription(ScriptContext context, PlanningItem plan, LocalScriptContext localContext)
        {
            List<Tuple<string, string, int, string, string, string, string, Tuple<string>>> constraintsList = new List<Tuple<string, string, int, string, string, string, string, Tuple<string>>>();
            SqlConnection conn = localContext.MySQLConnection.Conn;
            if (context.PlanSetup != null)
            {
                string planUID = ((PlanSetup)plan).UID;
                constraintsList = GetFullRTPrescriptionDataBase(planUID, localContext);
            }
            else if (context.PlanSumsInScope.Count() > 0)// Atenció només accepta el "primer" PlanSum
            {
                foreach (PlanSetup planS in ((PlanSum)plan).PlanSetups)
                {
                    string planUID = ((PlanSetup)planS).UID;

                    constraintsList = GetFullRTPrescriptionDataBase(planUID, localContext);

                    if (constraintsList.Any(t => t.Item3 == 3))
                    {
                        return constraintsList;
                    }
                }
            }
            return constraintsList;
        }

        /// <summary>
        /// Get a full RT Prescription for a specific plan (Physician's intent).
        /// - Prescription (VolumeId, dose per fraction, total dose)
        /// - Volume dose coverage (maximum dose, minimum dose, minimum dvh dose, maximum dvh dose)
        /// - Constraints (maximum mean dose, maximum dose, OAR FREE TEXT)
        /// </summary>
        /// <param name="planUID">
        /// The identifier of the Plan. 
        /// </param>
        /// <param name="localContext">
        /// Represent the local context
        /// </param>
        /// <returns>
        /// List of Tuples where a full RT Prescription is stored.
        /// Tuple -> ItemValue,ItemType,AnatomyRole,AnatomyName,IntemValueUnit,PrescriptionName,Param1Value,Param1ValueUnit
        /// For more information of each item in tuple see "data base information"
        /// </returns>
        public static List<Tuple<string, string, int, string, string, string, string, Tuple<string>>> GetFullRTPrescriptionDataBase(string planUID, LocalScriptContext localContext)
        {
            SqlDataReader reader;
            List<Tuple<string, string, int, string, string, string, string, Tuple<string>>> constraintsList = new List<Tuple<string, string, int, string, string, string, string, Tuple<string>>>();
            SqlConnection conn = localContext.MySQLConnection.Conn;
            try
            {
                string cmdStr = @"SELECT dbo.PrescriptionAnatomyItem.ItemValue,   
                                     dbo.PrescriptionAnatomyItem.ItemType,          
                                     dbo.PrescriptionAnatomy.AnatomyRole,   
                                     dbo.PrescriptionAnatomy.AnatomyName,   
                                     dbo.PrescriptionAnatomyItem.ItemValueUnit,   
                                     dbo.Prescription.PrescriptionName,   
                                     dbo.PrescriptionAnatomyItem.Param1Value,   
                                     dbo.PrescriptionAnatomyItem.Param1ValueUnit  
                                FROM dbo.PlanSetup,   
                                     dbo.Prescription,   
                                     dbo.PrescriptionAnatomy,   
                                     dbo.PrescriptionAnatomyItem,   
                                     dbo.RTPlan  
                               WHERE ( dbo.PrescriptionAnatomyItem.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer ) and  
                                     ( dbo.RTPlan.PlanSetupSer = dbo.PlanSetup.PlanSetupSer ) and  
                                     ( dbo.Prescription.PrescriptionSer = dbo.PlanSetup.PrescriptionSer ) and  
                                     ( dbo.PrescriptionAnatomy.PrescriptionSer = dbo.Prescription.PrescriptionSer ) and  
                                     ( ( dbo.RTPlan.PlanUID = @planUID ) ) ";

                using (SqlCommand command = new SqlCommand(cmdStr, conn))
                {
                    string ItemValue = string.Empty;
                    string ItemType = string.Empty;
                    int AnatomyRole = 0;
                    string AnatomyName = string.Empty;
                    string ItemValueUnit = string.Empty;
                    string PrescriptionName = string.Empty;
                    string Param1Value = string.Empty;
                    string Param1ValueUnit = string.Empty;

                    command.Parameters.AddWithValue("@planUID", planUID);
                    reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        try
                        {
                            ItemValue = reader.GetString(0);
                            ItemType = reader.GetString(1);
                            AnatomyRole = reader.GetInt16(2);
                            AnatomyName = reader.GetString(3);
                            if (!reader.IsDBNull(4)) { ItemValueUnit = reader.GetString(4); } else { ItemValueUnit = string.Empty; }
                            PrescriptionName = reader.GetString(5);
                            if (!reader.IsDBNull(6)) { Param1Value = reader.GetString(6); } else { Param1Value = string.Empty; }
                            if (!reader.IsDBNull(7)) { Param1ValueUnit = reader.GetString(7); } else { Param1ValueUnit = string.Empty; }
                            var constraintsTuple = Tuple.Create(ItemValue, ItemType, AnatomyRole, AnatomyName, ItemValueUnit, PrescriptionName, Param1Value, Param1ValueUnit);
                            constraintsList.Add(constraintsTuple);
                        }
                        catch (Exception)
                        {
                            reader.Close();
                        }
                    }
                }
                reader.Close();
                return constraintsList;
            }
            catch (Exception e)
            {
                throw new Exception("CheckRTPrescriptionExists Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get a basic RT Prescription for a specific plan (Physician's intent).
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
        public static List<Tuple<string, double, double>> GetBasicRTPrescription(PlanSetup plan, LocalScriptContext localContext)
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
                                 ( ( dbo.RTPlan.PlanUID = @planUID ) ) "; // AnatomyRole = 2 means RTPrescription

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
    }

    #endregion

    #region CONSTRAINTS
    public class Constraints
    {
        //(Item1)ItemValue,(Item2)ItemType,(Item3)AnatomyRole,(Item4)AnatomyName,(Item5)IntemValueUnit,
        //(Item6)PrescriptionName,(Item7)Param1Value,(Item8)Param1ValueUnit
        public static void CheckConstraints(List<Tuple<string, string, int, string, string, string, string, Tuple<string>>> rtPrescription, IEnumerable<Structure> myStructures, PlanningItem myPlan, Fractionation myPrescription)
        {
            try
            {
                // ------- Check if each structure has a RT Prescription . If not, the error message is written. --------
                string structureNoRTPrescription = string.Empty;
                
                foreach (Structure structure in myStructures)
                {
                    bool exists = false;
                    string structureID = Regex.Replace(StringExtensions.RemoveDiacritics(structure.Id), @"\s", "").ToLower(); // remove accents, remove spaces, convert to lowercase.     
                    string anatomyName = string.Empty;

                    foreach (Tuple<string, string, int, string, string, string, string, Tuple<string>> t in rtPrescription)
                    {
                        anatomyName = Regex.Replace(StringExtensions.RemoveDiacritics(t.Item4), @"\s", "").ToLower();
                        if (anatomyName.Like(structureID) && t.Item3 == 3) //Item3 == 3 means constraint (neither prescription nor volume coverage)
                        {
                            exists = true;
                            break;
                        }
                    }
                    
                    if (!exists && (!((Regex.Match(structureID, @"^[pcg]TV", RegexOptions.IgnoreCase).Success) ||  // Regex.Match(structureID, @"^[pcg]TV", RegexOptions.IgnoreCase -> empieza por "p", "c", o "g" (^[pcg]) seguido de "tv" (TV) y después cualquier carácter.
                        (Regex.Match(structureID, @"drr", RegexOptions.IgnoreCase).Success) || 
                        (Regex.Match(structureID, @"fis", RegexOptions.IgnoreCase).Success) || 
                        (Regex.Match(structureID, @"couch", RegexOptions.IgnoreCase).Success) || 
                        structure.DicomType == "EXTERNAL"))) // EXTERNAL = Body.
                    {
                        structureNoRTPrescription += ("\t" + structure.Id + "\n");
                    }
                }

                // ------- Check if each constraint of RT prescrition has a Structure. If not, the error message is written.
                // If constraint has a Structure, the comparison with plan is made.  -------- 

                string constraintNoStructure = string.Empty;
                List<string> volumes = new List<string>(); // Auxiliar volume list (not repeat volumes)

                foreach (var localRTPrescription in rtPrescription)
                {
                    string anatomyName = Regex.Replace(StringExtensions.RemoveDiacritics(localRTPrescription.Item4), @"\s", "").ToLower();  // remove accents, remove spaces, convert to lowercase. 
                    int anatomyRole = localRTPrescription.Item3; 

                    foreach (Structure structure in myStructures)
                    {                        
                        string structureID = Regex.Replace(StringExtensions.RemoveDiacritics(structure.Id), @"\s", "").ToLower();  
                                                                  
                        if (Regex.Replace(structureID, @"\s", "") == anatomyName && anatomyRole == 3 &&  //Item3 == 3 means constraint (neither prescription nor volume coverage)
                        (!((Regex.Match(structureID, @"^[pcg]TV", RegexOptions.IgnoreCase).Success) ||  // Regex.Match(structureID, @"^[pcg]TV", RegexOptions.IgnoreCase -> empieza por "p", "c", o "g" (^[pcg]) seguido de "tv" (TV) y después cualquier carácter.
                        (Regex.Match(structureID, @"drr", RegexOptions.IgnoreCase).Success) ||
                        (Regex.Match(structureID, @"fis", RegexOptions.IgnoreCase).Success) ||
                        (Regex.Match(structureID, @"couch", RegexOptions.IgnoreCase).Success) ||
                        structure.DicomType == "EXTERNAL")))  // EXTERNAL = Body.
                        {
                            bool objectiveAchievedBool = false;

                            // Check if localRTPrescription.Item1 has format "numbers + %" (Example. 25%) or only numbers (Example. 25); 
                            Regex regex = new Regex(@"^\d+%$"); //Empezar por (^) uno o mas números (\d+) y le sigue y finaliza con el símbolo % (%$). 
                            Regex regex1 = new Regex(@"^\d+$"); //Empezar por (^) uno o mas números (\d+) y que finalice con número($).
                            Match match = regex.Match(Regex.Replace(localRTPrescription.Item1, @"\s$", "", RegexOptions.IgnoreCase).ToLower()); //\s$ -> eliminar los espacios finales si los hay
                            Match match1 = regex1.Match(Regex.Replace(localRTPrescription.Item1, @"\s$", "", RegexOptions.IgnoreCase).ToLower());

                            if (match.Success || match1.Success)
                            {
                                if (localRTPrescription.Item2.Like("MAXIMUM MEAN DOSE") || localRTPrescription.Item2.Like("MAXIMUM DOSE"))
                                {
                                    string RTPrescriptionObjectivePrint = localRTPrescription.Item1.ToString() + " Gy";

                                    double objectiveAchievedValue = double.NaN;
                                    string objectiveAchievedValuePrint = string.Empty;
                                    if (localRTPrescription.Item2 == "MAXIMUM MEAN DOSE")
                                    {                                        
                                        objectiveAchievedValue = CheckDoseVolume.GetMeanDose(myPlan, "Absolute", structure).Dose;
                                        objectiveAchievedValuePrint = "Dmean = " + Math.Round(objectiveAchievedValue, 1) + " Gy";                                        
                                        objectiveAchievedBool = (Double.Parse(localRTPrescription.Item1) > objectiveAchievedValue);
                                    }
                                    else if (localRTPrescription.Item2 == "MAXIMUM DOSE")
                                    {
                                        objectiveAchievedValue = CheckDoseVolume.GetMaxDose(myPlan, "Absolute", structure).Dose;
                                        objectiveAchievedValuePrint = "Dmax = " + Math.Round(objectiveAchievedValue, 1) + " Gy";
                                        objectiveAchievedBool = (Double.Parse(localRTPrescription.Item1) > objectiveAchievedValue);
                                    }

                                    Script.constraintsAnalyzed.Rows.Add(structure.Id, localRTPrescription.Item2, RTPrescriptionObjectivePrint, objectiveAchievedValuePrint, objectiveAchievedBool);
                                }
                                else // localRTPrescription.Item2 = "OAR FREE TEXT"
                                {
                                   
                                    // V20 < 25% or V20 < 1cm3 <---> doseValue < volumeValue(VolumePresentation)
                                    double volumeValue = CheckDoseVolume.GetVolumeValue(localRTPrescription.Item1);
                                    string volumePresentation = CheckDoseVolume.GetVolumePresentation(localRTPrescription.Item1);
                                    double doseValue = CheckDoseVolume.GetDoseValue(localRTPrescription.Item7);
                                    string RTPrescriptionObjectivePrint = string.Empty;
                                    double objectiveAchievedValue = double.NaN;
                                    string objectiveArchiveValuePrint = string.Empty;
                                    double objectiveAchievedPercentVolum = double.NaN;
                                    objectiveAchievedBool = false;

                                    if (doseValue != -1) // <-- Check if the dose value has a correct format "only numbers" or "V+numbers" (V20 means Volume which receives 20 Gy)
                                    {
                                        if (volumePresentation == "AbsoluteCm3") // ¡¡¡¡¡¡¡¡¡¡¡ AbsoluteCm3 NOT VALIDATED !!!!!!!!!!!!!!!!!!
                                        {
                                            RTPrescriptionObjectivePrint = ("V" + Math.Round((double)doseValue, 1).ToString() + "Gy" + " < " + volumeValue + " cm3");
                                            //objectiveAchievedPercentVolum = CheckDoseVolume.GetPercentVolume(myPlan, structure, doseValue, "Absolute");
                                            objectiveAchievedValue = CheckDoseVolume.GetAbsoluteDose(myPlan, structure, volumeValue, VolumePresentation.AbsoluteCm3).Dose; // Absolute Volume
                                            objectiveArchiveValuePrint = ("V" + Math.Round((double)objectiveAchievedValue, 1).ToString() + "Gy" + " = " + volumeValue + " cm3");
                                            //objectiveArchiveValuePrint = ("V" + doseValue.ToString() + " = " + objectiveAchievedPercentVolum.ToString() + "%");
                                            //objectiveAchievedBool = (objectiveAchievedValue < doseValue);
                                            objectiveAchievedBool = (objectiveAchievedValue < volumeValue);
                                        }
                                        else if (volumePresentation == "Relative")
                                        {
                                            RTPrescriptionObjectivePrint = ("V" + Math.Round((double)doseValue, 1).ToString() + "Gy" + " < " + volumeValue + " %");
                                            objectiveAchievedPercentVolum = CheckDoseVolume.GetPercentVolume(myPlan, structure, doseValue, "Absolute", myPrescription); // -----> It has the PlanSum into account.
                                            objectiveAchievedValue = CheckDoseVolume.GetAbsoluteDose(myPlan, structure, volumeValue, VolumePresentation.Relative).Dose; //Relative Volume
                                            objectiveArchiveValuePrint = ("V" + Math.Round((double)doseValue, 1).ToString() + "Gy" + " = " + Math.Round((double)objectiveAchievedPercentVolum, 1).ToString() + "%");
                                            objectiveAchievedBool = (objectiveAchievedValue < doseValue);
                                        }
                                        else
                                        {
                                            throw new SystemException();
                                        }

                                        Script.constraintsAnalyzed.Rows.Add(structure.Id, "DOSE VOLUME", RTPrescriptionObjectivePrint, objectiveArchiveValuePrint, objectiveAchievedBool);
                                    }
                                }
                            }
                        }
                        else if (anatomyRole == 3)
                        {                           
                            bool structureIDExists = myStructures.Any(item => Regex.Replace(StringExtensions.RemoveDiacritics(item.Id), @"\s", "").ToLower() == anatomyName);
                            if (!structureIDExists && !volumes.Exists(t => t == anatomyName))                           
                            {
                                constraintNoStructure += ("\t" + localRTPrescription.Item4 + "\n"); // not anatomyName because we want to show the original name.
                                volumes.Add(anatomyName);
                            }
                        }
                    }
                }

                if (structureNoRTPrescription != string.Empty)
                {
                    Script.constraintsErrors.Rows.Add("There are structure(s) without constraint: \n" + structureNoRTPrescription);
                }

                if (constraintNoStructure != string.Empty)
                {
                    Script.constraintsErrors.Rows.Add("The following constraints do not match any structure: \n" + constraintNoStructure);
                }

            }
            catch (Exception e)
            {
                throw new Exception("CheckConstraints function Error: \n" + e.Message +"    "+e.StackTrace);
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
        /// Get Mean Structure Dose.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="presentation">
        /// Dose and volume presentation in DVH (Absolute or Relative)
        /// </param>
        /// <param name="str">
        /// Structure Object
        /// </param>
        /// <returns> Mean Structe Dose </returns>
        public static DoseValue GetMeanDose(PlanningItem plan, string presentation, Structure str)
        {
            try
            {
                DVHData dvhData = null;

                if (String.Equals(presentation, "relative", StringComparison.OrdinalIgnoreCase))
                {
                    dvhData = plan.GetDVHCumulativeData(str,
                        DoseValuePresentation.Relative,
                        VolumePresentation.Relative, 0.1);
                }
                else if (String.Equals(presentation, "absolute", StringComparison.OrdinalIgnoreCase))
                {
                    dvhData = plan.GetDVHCumulativeData(str,
                      DoseValuePresentation.Absolute,
                      VolumePresentation.AbsoluteCm3, 0.1);
                }
                DoseValue meanDose = dvhData.MeanDose;
                return meanDose;
            }
            catch (Exception e)
            {
                throw new Exception("GetMeanDose function Error: \n       Possible Target Volume not Valid \n" + e.Message);
            }

        }

        /// <summary>
        /// Get Maximum Structure Dose.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="presentation">
        /// Dose and volume presentation in DVH (Absolute or Relative)
        /// </param>
        /// <param name="str">
        /// Structure Object
        /// </param>
        /// <returns> Maximum Structe Dose </returns>
        public static DoseValue GetMaxDose(PlanningItem plan, string presentation, Structure str)
        {
            try
            {
                DVHData dvhData = null;

                if (String.Equals(presentation, "relative", StringComparison.OrdinalIgnoreCase))
                {
                    dvhData = plan.GetDVHCumulativeData(str,
                        DoseValuePresentation.Relative,
                        VolumePresentation.Relative, 0.1);
                }
                else if (String.Equals(presentation, "absolute", StringComparison.OrdinalIgnoreCase))
                {
                    dvhData = plan.GetDVHCumulativeData(str,
                      DoseValuePresentation.Absolute,
                      VolumePresentation.AbsoluteCm3, 0.1);
                }
                DoseValue maxDose = dvhData.MaxDose;
                return maxDose;
            }
            catch (Exception e)
            {
                throw new Exception("GetMaxDose function Error: \n       Possible Target Volume not Valid \n" + e.Message);
            }

        }

        /// <summary>
        /// Get Maximum Structure Dose.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="presentation">
        /// Dose and volume presentation in DVH (Absolute or Relative)
        /// </param>
        /// <param name="str">
        /// Structure Object
        /// </param>
        /// <returns> Maximum Structe Dose </returns>
        public static DoseValue GetMinDose(PlanningItem plan, string presentation, Structure str)
        {
            try
            {
                DVHData dvhData = null;

                if (String.Equals(presentation, "relative", StringComparison.OrdinalIgnoreCase))
                {
                    dvhData = plan.GetDVHCumulativeData(str,
                        DoseValuePresentation.Relative,
                        VolumePresentation.Relative, 0.1);
                }
                else if (String.Equals(presentation, "absolute", StringComparison.OrdinalIgnoreCase))
                {
                    dvhData = plan.GetDVHCumulativeData(str,
                      DoseValuePresentation.Absolute,
                      VolumePresentation.AbsoluteCm3, 0.1);
                }
                DoseValue minDose = dvhData.MinDose;
                return minDose;
            }
            catch (Exception e)
            {
                throw new Exception("GetMinDose function Error: \n       Possible Target Volume not Valid \n" + e.Message);
            }

        }

        /// <summary>
        /// Get the percentage of volume which recive a specific dose.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="str">
        /// Structure object.
        /// </param>
        /// <param name="value">
        /// Absolute or percentage dose value
        /// </param>
        /// <param name="presentation">
        /// "Absolute" -> Absolute dose presentation
        /// "Relative" -> Relative dose presentation
        /// </param>
        /// <returns>
        /// Percentage of volume.
        /// </returns>
        public static double GetPercentVolume(PlanSetup plan, Structure str, double value, string presentation) 
        {
            try
            {
                DoseValue DV;
                if (String.Equals(presentation, "relative", StringComparison.OrdinalIgnoreCase))
                {
                    DV = new DoseValue(value, "%");
                }
                else if (String.Equals(presentation, "absolute", StringComparison.OrdinalIgnoreCase))
                {
                    DV = new DoseValue(value, "Gy");
                }
                else
                {
                    throw new Exception("Presentation parameter is not 'Relative' or 'Absolute'");
                }
                
                double volumePercent = plan.GetVolumeAtDose(str, DV, VolumePresentation.Relative);
                if (volumePercent.ToString().Like("NeuN"))
                {
                    volumePercent = 0;
                }
                return volumePercent;
            
            }
            catch (Exception e)
            {
                throw new Exception("GetPercentVolume function Error: \n" + e.Message);
            }

        }

        
        /// <summary>
        /// Get the percentage of volume which recive a specific dose.
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset or plan sum dataset
        /// </param>
        /// <param name="str">
        /// Structure object.
        /// </param>
        /// <param name="value">
        /// Absolute or percentage dose value
        /// </param>
        /// <param name="presentation">
        /// "Absolute" -> Absolute dose presentation
        /// "Relative" -> Relative dose presentation
        /// </param>
        /// <param name="myPrescription">
        /// Prescription plan or prescription plan sum
        /// </param>
        /// <returns>
        /// Percentage of volume.
        /// </returns>
        public static double GetPercentVolume(PlanningItem plan, Structure str, double value, string presentation, Fractionation myPrescription)
        {
            try
            {
                DoseValue DV;
                double volumePercent = double.NaN;
                if (String.Equals(presentation, "Relative", StringComparison.OrdinalIgnoreCase))
                {
                    DV = new DoseValue(value, "%");
                }
                else if (String.Equals(presentation, "Absolute", StringComparison.OrdinalIgnoreCase))
                {
                    DV = new DoseValue(value, "Gy");
                }
                else
                {
                    throw new Exception("Presentation parameter is not 'Relative' or 'Absolute'");
                }

                if (plan is PlanSetup)
                {
                    volumePercent = ((PlanSetup)plan).GetVolumeAtDose(str, DV, VolumePresentation.Relative);                    
                }
                else
                {
                    if (DV.UnitAsString == "percent") //Convert the percent in absolute. Mandatory in Plan Sum.
                    {
                        double totalAbsoluteDose = (double)myPrescription.NumberOfFractions * myPrescription.PrescribedDosePerFraction.Dose;
                        double absoluteDose = totalAbsoluteDose * DV.Dose / 100;
                        DV = new DoseValue(absoluteDose, "Gy");
                    }
                    DVHData myDVH = plan.GetDVHCumulativeData(str, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.001);
                    DVHPoint[] myDVHpoints = myDVH.CurveData;
                    foreach (DVHPoint p in myDVHpoints)
                    {
                        if (p.DoseValue.Dose > DV.Dose)
                        {
                            volumePercent = p.Volume;
                            break;
                        }
                    }
                    
                }              
                  
                if (volumePercent.ToString().Like("NeuN"))
                {
                    volumePercent = 0;
                }
                return volumePercent;
            }
            catch (Exception e)
            {
                throw new Exception("GetPercentVolume function Error: \n" + e.Message);
            }

        }

        /// <summary>
        /// Get the absolute dose which covers a specific percentage or absolute volume
        /// </summary>
        /// <param name="plan">
        /// Represents a treatment plan dataset.
        /// </param>
        /// <param name="str">
        /// Structure object.
        /// </param>
        /// <param name="volum">
        /// Absolute or percentage volum value
        /// </param>
        /// <param name="volumPresentation">
        /// "AbsoluteCm3" -> Absolute volume presentation
        /// "Relative" -> Relative volume presentation
        /// </param>
        /// <returns>
        /// Absolute dose
        /// </returns>
        public static DoseValue GetAbsoluteDose(PlanningItem plan, Structure str, double volume, VolumePresentation volumePresentation)
        {

            
            DoseValue dose = DoseValue.UndefinedDose();
            try
            {
                if (plan is PlanSetup)
                {
                    if (String.Equals(volumePresentation.ToString(), "Relative", StringComparison.OrdinalIgnoreCase))
                    {
                        //dose = plan.GetDoseAtVolume(str, volum, VolumePresentation.Relative, DoseValuePresentation.Absolute);
                        dose = ((PlanSetup)plan).GetDoseAtVolume(str, volume, VolumePresentation.Relative, DoseValuePresentation.Absolute);
                        return dose;
                    }
                    else if (String.Equals(volumePresentation.ToString(), "AbsoluteCm3", StringComparison.OrdinalIgnoreCase))
                    {
                        dose = ((PlanSetup)plan).GetDoseAtVolume(str, volume, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute);
                        return dose;
                    }
                    else
                    {
                        throw new Exception("Dose presentation parameter is not 'Relative' or 'AbsoluteCm3'");
                    }
                }
                else
                {
                    DVHPoint[] myDVHpoints = plan.GetDVHCumulativeData(str, DoseValuePresentation.Absolute, volumePresentation, 0.001).CurveData;
                    foreach (DVHPoint p in myDVHpoints)
                    {
                        if (p.Volume < volume)
                        {
                            dose = p.DoseValue;
                            break;
                        }
                    }
                    return dose;
                }
                    
            }
            catch (Exception e)
            {
                throw new Exception("GetAbsoluteDose function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get volume presentation.
        /// </summary>
        /// <param name="s">
        /// String which represents a volume in percentage (25%) or cubic centimeters (10cc or 10cm3)
        /// </param>
        /// <returns>
        /// String "AbsoluteCm3" if volume -> cubic centimeters or "Relative" if volume -> percentage.
        /// </returns>
        public static string GetVolumePresentation (string s)
        {
            try
            {
                string volumPresentation = string.Empty;
                if (s.IndexOf("cc") != -1 || s.IndexOf("cm3") != -1)
                {
                    volumPresentation = "AbsoluteCm3";
                }
                else
                {
                    volumPresentation = "Relative";
                }
                return volumPresentation;
            }
            catch (Exception e)
            {
                throw new Exception("GetVolumePresentation function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get volume value
        /// </summary>
        /// <param name="s">
        /// String which represents a volume in percentage (25%) or cubic centimeters (10cc or 10cm3)
        /// </param>
        /// <returns>
        /// Volume value without unit of measurement.
        /// </returns>
        public static int GetVolumeValue (string s)
        {
            try
            {
                string volumValue = Regex.Match(s, @"\d+").Value;
                return Int32.Parse(volumValue);
            }
            catch (Exception e)
            {
                throw new Exception("GetVolumeValue function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Get a dose value
        /// </summary>
        /// <param name="s">
        /// String which represents the dose received by a volume. (V20 or 20 -> 20 Gy)
        /// </param>
        /// <returns>
        /// Double value which represents the dose.
        /// </returns>
        public static double GetDoseValue (string s)
        {
            try
            {
                if (Regex.Match(s, @"\d+", RegexOptions.IgnoreCase).Success || Regex.Match(s, @"^V\d", RegexOptions.IgnoreCase).Success) // Format "only numbers" or "V+numbers"
                {
                    string doseValue = Regex.Match(s, @"\d+").Value;
                    return Double.Parse(doseValue);
                }
                else
                {
                    return -1.0;
                }
            }
            catch (Exception e)
            {
                throw new Exception("GetDoseValue function Error: \n" + e.Message);
            }
        }
    }

    #endregion

    #region HISTOGRAM MANAGEMENT
    //public static class ExtensionMethods
    //{

    //    public static DoseValue GetDoseAtVolume(this PlanningItem myPlan, Structure structure, double volume, VolumePresentation volumePresentation, DoseValuePresentation dosePresentation)
    //    {
    //        DoseValue myDose = DoseValue.UndefinedDose();
    //        if (myPlan is PlanSetup)
    //            return ((PlanSetup)myPlan).GetDoseAtVolume(structure, volume, volumePresentation, dosePresentation);
    //        else
    //        {
    //            DVHPoint[] myDVHpoints = myPlan.GetDVHCumulativeData(structure, DoseValuePresentation.Absolute, volumePresentation, 0.001).CurveData;
    //            foreach (DVHPoint p in myDVHpoints)
    //            {
    //                if (p.Volume < volume)
    //                {
    //                    myDose = p.DoseValue;
    //                    break;
    //                }
    //            }
    //            return myDose;
    //        }
    //    }

    //    public static double GetVolumeAtDose(this PlanningItem myPlan, Structure structure, DoseValue dose, VolumePresentation volumePresentation)
    //    {
    //        double myVolume = 0.0;
    //        if (myPlan is PlanSetup)
    //            return ((PlanSetup)myPlan).GetVolumeAtDose(structure, dose, volumePresentation);

    //        else
    //        {
    //            DVHData myDVH = myPlan.GetDVHCumulativeData(structure, DoseValuePresentation.Absolute, volumePresentation, 0.001);
    //            DVHPoint[] myDVHpoints = myDVH.CurveData;
    //            foreach (DVHPoint p in myDVHpoints)
    //            {
    //                if (p.DoseValue.Dose > dose.Dose)
    //                {
    //                    myVolume = p.Volume;
    //                    break;
    //                }
    //            }
    //            return myVolume;
    //        }
    //    }
    //}
    #endregion

    #region VERIFICATIONS
    public class Verifications
    {
        /// <summary>
        /// Check initial parameters to run the Script.       
        /// </summary> 
        /// <param name="context">
        /// The TPS context for a Plan or Plan Sum
        /// </param>
        /// <param name="msg">
        /// Out message with errors.
        /// </param>
        /// <returns>
        /// True if some error exists
        /// </returns>
        public static bool CheckInitialParameters(ScriptContext context, out string msg)
        {

            msg = string.Empty;
            bool error = false;
            PlanningItem myPlan;

            try
            {
                if (context.PlanSetup != null)
                {
                    myPlan = context.PlanSetup;
                    if (Verifications.CheckIndivualParameters((PlanSetup)myPlan, out msg))
                    {
                        msg += msg;
                        error = true;
                    }
                }
                else if (context.PlanSumsInScope.Count() > 0)// Atenció només accepta el "primer" PlanSum
                {
                    myPlan = context.PlanSumsInScope.FirstOrDefault();
                    foreach (PlanSetup planS in ((PlanSum)myPlan).PlanSetups)
                    {
                        if (Verifications.CheckIndivualParameters(planS, out msg))
                        {
                            msg += ("Plan errors in " + planS.Id + "\n");
                            msg += msg;
                            error = true;
                        }
                    }
                }
                else
                {
                    throw new Exception("Neither Plan nor PlanSum has not been correctly selected.");
                }

                return error;

            }
            catch (Exception e)
            {
                throw new Exception("CheckInitialParameters function Error: \n" + e.Message);
            }
        }
        /// <summary>
        /// Check initial parameters to run the Script.       
        /// </summary> 
        /// <param name="plan">
        /// Represents a plan (not valid plan sum)
        /// </param>
        /// <param name="msg">
        /// Message with errors
        /// </param>
        /// <returns>
        /// True if some error exists
        /// </returns>
        public static bool CheckIndivualParameters(PlanSetup plan, out string msg)
        {
            try
            {
                bool error = false;
                msg = string.Empty;
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
                        CheckPrescription.GetNumberOfFracctions(plan);
                        CheckPrescription.GetDosePerFracction(plan);
                        CheckPrescription.GetTotalDose(plan);
                    }
                    catch (Exception)
                    {
                        msg += ("\u2022   Check the Plan Dose prescription.\n");
                        error = true;
                    }
                }

                return error;
            }
            catch (Exception e)
            {
                throw new Exception("CheckIndivualParameters function Error: \n" + e.Message);
            }
        }
    }
    #endregion

    #region PLAN PRESCRIPTION FUNCTIONS
    /// <summary>
    /// Plan Presctiption Functions
    /// </summary>
    public class CheckPrescription
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
    }
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

    #region PRINTING FUNCTIONS

    /// <summary>
    /// Printing functions
    /// </summary>
    public class CheckPrinting
    {

        /// <summary>
        /// Print the constraint error messages table.
        /// </summary>
        /// <param name="constraintsWarningsTable">
        /// Table with constraint warnings
        /// </param>
        public static void PrintConstraintWarnings(DataTable constraintsWarningsTable)

        {
            try
            {
                string printConstraintsMessage = string.Empty;               

                if (constraintsWarningsTable != null && constraintsWarningsTable.Rows.Count > 0)
                {
                    foreach (DataRow row in constraintsWarningsTable.Rows)
                    {
                        printConstraintsMessage += ("\u2022    " + row["ConstraintErrors"] + "\n");
                    }

                    MessageBox.Show(printConstraintsMessage, "Warnings");
                }
                else
                {
                    MessageBox.Show("\n\n No Warnings\n\n", "Warnings");
                }
            }
            catch (Exception e)
            {
                throw new Exception("PrintConstraintWarnings function Error: \n" + e.Message);
            }
        }


        /// <summary>
        /// Print the analyzed constraints messages table.
        /// </summary>
        /// <param name="constraintsAnalyzedTable">
        /// Table with the constraints analyzed
        /// </param>
        public static void PrintConstraintsAnalyzed(DataTable constraintsAnalyzedTable, PlanningItem myPlan)
        {
            try
            {
                if (constraintsAnalyzedTable != null && constraintsAnalyzedTable.Rows.Count > 0)
                {
                    int dataRowCount = 0;
                    int trueCount = 0;
                    foreach (DataRow row in constraintsAnalyzedTable.Rows)
                    {
                        dataRowCount++;
                        if ((bool)row["ObjectiveAchievedBool"] == true)
                        {
                            trueCount++;
                        }                     
                    }

                    string printVerificationMessage = string.Empty;
                    string RTPrescriptionObjectiveLabel = string.Empty;
                    string currentStructureId = string.Empty;
                    string lastStructureId = string.Empty;


                    printVerificationMessage += (myPlan.Id + " :" +"\n");  // Print the plan ID

                    if (dataRowCount == trueCount) // Every line have the OK symbol
                    {
                        foreach (DataRow row in constraintsAnalyzedTable.Rows)
                        {
                            //Insert line between different OAR
                            currentStructureId = row["ORName"].ToString();
                            if (!currentStructureId.Like(lastStructureId))
                            {
                                printVerificationMessage += "\n";
                                lastStructureId = currentStructureId;
                            }

                            // Print verification message line                           
                            printVerificationMessage += CheckPrinting.PrintVerificationMessage("\u2714",row); //"\u2714" is the "ok" symbol in unicode

                        }
                    }
                    else // Every line where the constraint has failed, the "x" symbol is inserted. The rest of lines nothing.
                    {
                        foreach (DataRow row in constraintsAnalyzedTable.Rows)
                        {
                            //Insert line between different OAR
                            currentStructureId = row["ORName"].ToString();
                            if (!currentStructureId.Like(lastStructureId))
                            {
                                printVerificationMessage += "\n";
                                lastStructureId = currentStructureId;
                            }

                            // Print verification message line                            
                            if ((bool)row["ObjectiveAchievedBool"] == false )
                            {
                                printVerificationMessage += CheckPrinting.PrintVerificationMessage("\u2718", row); //"\u2718" is the "x" symbol in unicode
                            }
                            else
                            {
                                printVerificationMessage += CheckPrinting.PrintVerificationMessage("   ", row); //"\u0020" is the "space" symbol in unicode

                            }
                        }
                    }

                    MessageBox.Show(printVerificationMessage, "Verification");
                }
            }
            catch (Exception e)
            {
                throw new Exception("PrintConstraintsAnalyzed function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// designs the verification row depends on type of constraint
        /// </summary>
        /// <param name="simbol">
        /// Simbol which represents spaces, ok or bad
        /// </param>
        /// <param name="row">
        ///  row in constraintsAnalyzedTable
        /// </param>
        /// <returns>
        /// New dessigned verification row. 
        /// </returns>
        public static string PrintVerificationMessage(string simbol, DataRow row)
        {
            try {
                string RTPrescriptionObjectiveLabel = string.Empty;
                string printVerificationMessage = string.Empty;

                if (row["RTPrescriptionObjectiveLabel"].ToString().Equals("MAXIMUM MEAN DOSE"))
                {
                    RTPrescriptionObjectiveLabel = "Dmean < ";
                }
                else if (row["RTPrescriptionObjectiveLabel"].ToString().Equals("MAXIMUM DOSE"))
                {
                    RTPrescriptionObjectiveLabel = "Dmax < ";
                }
                else
                {
                    RTPrescriptionObjectiveLabel = "";
                }

                if (row["ORName"].ToString().Length > 5)
                {
                    printVerificationMessage +=
                        (simbol + "     " + row["ORName"] + "\t" + RTPrescriptionObjectiveLabel + row["RTPrescriptionObjectivePrint"] + "\t" + row["ObjectiveAchievedValue"] + "\n");  
                }
                else
                {
                    printVerificationMessage +=
                        (simbol + "     " + row["ORName"] + "\t\t" + RTPrescriptionObjectiveLabel + row["RTPrescriptionObjectivePrint"] + "\t" + row["ObjectiveAchievedValue"] + "\n");  
                }

                return printVerificationMessage;
            }
            catch (Exception e)
            {
                throw new Exception("PrintVerificationMessage function Error: \n" + e.Message);
            }
        }

    }
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
            try
            {
                return new Regex(
                    "^" + Regex.Escape(pattern).Replace(@"\*", ".*").Replace(@"\?", ".") + "$",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline
                ).IsMatch(str);
            }
            catch (Exception e)
            {
                throw new Exception("Like function Error: \n" + e.Message);
            }
        }

        /// <summary>
        /// Removes a text diacritics
        /// </summary>
        /// <param name="text">
        /// Script which represent the text.
        /// </param>
        /// <returns>
        /// Text without diacritics.
        /// </returns>
        public static string RemoveDiacritics(string text)
        {
            string formD = text.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            foreach (char ch in formD)
            {
                UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (uc != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(ch);
                }
            }

            return sb.ToString().Normalize(NormalizationForm.FormC);
        }
    }
    #endregion
}
