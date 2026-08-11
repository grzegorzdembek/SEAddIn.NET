using SolidEdgeAdd_In.Ribbons;

namespace SolidEdgeAdd_In
{
    /*
    The Solid Edge API provides an easy-to-use set of interfaces that enable programmers to fully integrate
    custom commands with Solid Edge. These custom programs are commonly referred to as addins.
    Specifically, Solid Edge defines an addin as a dynamically linked library (DLL) containing a COM-based
    object that implements the ISolidEdgeAddIn interface.
    Generally, an add-in is a COM object that is used to provide commands or other value to Solid Edge
    */

    // Define a COM-visible class with a GUID and ProgID for Solid Edge to identify this add-in
    [Guid("384E5BCF-DD43-49AC-BF01-F99CC009B35F"), ProgId("SolidEdgeAdd-In.Addin"), ComVisible(true)]
    public class AddIn : SeISolidEdgeAddIn
    {
        public SeAddIn m_addin; // Reference to the AddIn object     
        public SeApp m_application; // Reference to the Application object

        private RibbonController m_Controller; // Controller managing all ribbons for different environments

        public SeApp Application => m_application; // Allow other classes to access the Application (read-only property)

        #region ISolidEdgeAddIn Members   

        /*
        AddIn needs to do few needs when OnConnection is called:
        1. Connect to any Solid Edge application event sets the add-in plans on using by providing the appropriate sinks to the application object.
        2. Connect to the Solid Edge Add-in object's event set if the add-in plans to add any commands to any environments.
        3. Set the GUI version property of the Solid Edge Add-in object. 
        */

        public void OnConnection(object Application, SeConnectMode ConnectMode, SeAddIn AddInInstance)
        {
            // Store the COM references for later use
            m_addin = AddInInstance;
            m_application = (SeApp)Application;

            // Initialize ribbon controller to manage environment-specific ribbons
            m_Controller = new RibbonController(new SolidEdgeAddInWrapper(this));

            // Set Addin's GUI Version 
            AddInInstance.GuiVersion = 1;
        }

        /*
        AddIn needs to do few needs when OnConnectToEnviroment is called:
        1. Determine which environment is being connected by converting the category ID passed in.
        2. Add any environment-specific UI elements (ribbons, buttons) if this is the first time the add-in is loaded into this environment.
        3. Ensure any commands or buttons previously added are correctly initialized and persisted by Solid Edge for future sessions.
        */

        public void OnConnectToEnvironment(string EnvCatID, object pEnvironmentDispatch, bool bFirstTime)
        {
            Guid environmentCategory = new Guid(EnvCatID); // Convert string category ID to Guid to identify environment

            Guid CATID_SE_Draft = new("{08244193-B78D-11D2-9216-00C04F79BE98}"); // Draft environment GUID

            Guid CATID_SE_Part = new("{26618396-09D6-11d1-BA07-080036230602}"); // Part environment GUID
            Guid CATID_SE_SyncPart = new("{D9B0BB85-3A6C-4086-A0BB-88A1AAD57A58}"); // Sync Part enviroment GUID

            Guid CATID_SE_SM = new("{26618398-09D6-11D1-BA07-080036230602}"); // SheetMetal environment GUID
            Guid CATID_SE_SyncSM = new("{9CBF2809-FF80-4dbc-98F2-B82DABF3530F}"); // Sync SheetMetal environment GUID

            Guid CATID_SE_Assembly = new("{26618395-09D6-11d1-BA07-080036230602}"); // Assembly environment GUID

            Guid[] supportedEnvironments = new Guid[]
            {
                CATID_SE_Draft,
                CATID_SE_Part,
                CATID_SE_SyncPart,
                CATID_SE_SM,
                CATID_SE_SyncSM,
                CATID_SE_Assembly
            };

            if (supportedEnvironments.Contains(environmentCategory))
            {
                var ribbon = new SolidEdgeRibbon(m_application);
                m_Controller.Add(ribbon, environmentCategory, bFirstTime);
            }
        }

        /*
        AddIn needs to do few needs when OnDisconnection is called:
        1. Disconnect from any Solid Edge event sets it previously connected to.
        2. Disconnect from the Add-in event set, if any connections exist.
        3. Release any COM objects or interfaces obtained from the Solid Edge application.
        4. Close any storage or streams opened in the application's documents.
        5. Perform any additional cleanup, such as freeing allocated resources, to avoid memory leaks.
        */

        public void OnDisconnection(SeDisconnectMode DisconnectMode)
        {
            m_Controller?.Dispose();
            m_Controller = null;

            if (m_addin != null)
            {
                Marshal.ReleaseComObject(m_addin);
                m_addin = null;
            }

            if (m_application != null)
            {
                Marshal.ReleaseComObject(m_application);
                m_application = null;
            }
        }
        #endregion

        #region "Regasm.exe functions" 

        // Called by Regasm.exe 
        [ComRegisterFunction]
        static void RegisterServer(Type t)
        {
            // Zastosowanie using var załatwia automatyczne zamykanie klucza po zakończeniu rejestracji
            using var baseKey = Registry.ClassesRoot.CreateSubKey($@"CLSID\{{{t.GUID}}}");

            if (baseKey != null)
            {
                // Tell Solid Edge to automatically load your addin
                baseKey.SetValue("AutoConnect", 1);

                // Write title
                if (t.Assembly.IsDefined(typeof(AssemblyTitleAttribute), true))
                {
                    var titleAttribute = (AssemblyTitleAttribute)AssemblyTitleAttribute.GetCustomAttribute(t.Assembly, typeof(AssemblyTitleAttribute));
                    baseKey.SetValue("409", titleAttribute.Title);
                }

                // Write description
                if (t.Assembly.IsDefined(typeof(AssemblyDescriptionAttribute), true))
                {
                    var descriptionAttribute = (AssemblyDescriptionAttribute)AssemblyDescriptionAttribute.GetCustomAttribute(t.Assembly, typeof(AssemblyDescriptionAttribute));

                    // Bezpieczne utworzenie i automatyczne zwolnienie podklucza (załataliśmy ewentualny wyciek pamięci!)
                    using var summaryKey = baseKey.CreateSubKey("Summary");
                    summaryKey?.SetValue("409", descriptionAttribute.Description);
                }

                // Write required registry entries for a Solid Edge Addin                  
                baseKey.CreateSubKey(@"Implemented Categories\{26B1D2D1-2B03-11d2-B589-080036E8B802}"); // CATID_SolidEdgeAddIn 

                // And for enviroments 
                baseKey.CreateSubKey(@"Environment Categories\{26618394-09D6-11d1-BA07-080036230602}"); // CATID_SEApplication

                baseKey.CreateSubKey(@"Environment Categories\{08244193-B78D-11D2-9216-00C04F79BE98}"); // CATID_SE Draft

                baseKey.CreateSubKey(@"Environment Categories\{26618396-09D6-11d1-BA07-080036230602}"); // CATID_SE Part
                baseKey.CreateSubKey(@"Environment Categories\{D9B0BB85-3A6C-4086-A0BB-88A1AAD57A58}"); // CATID_SE Part SYNC

                baseKey.CreateSubKey(@"Environment Categories\{26618398-09D6-11D1-BA07-080036230602}"); // CATID_SE SM
                baseKey.CreateSubKey(@"Environment Categories\{9CBF2809-FF80-4dbc-98F2-B82DABF3530F}"); // CATID_SE SM SYNC

                baseKey.CreateSubKey(@"Environment Categories\{26618395-09D6-11d1-BA07-080036230602}"); // CATID_SE Assembly
            }
        }

        // Here we cleanup any registry values left from Regasm /u.
        [ComUnregisterFunction]
        static void UnregisterServer(Type t)
        {
            Registry.ClassesRoot.DeleteSubKeyTree($@"CLSID\{{{t.GUID}}}", false);
        }
        #endregion
    }
}

