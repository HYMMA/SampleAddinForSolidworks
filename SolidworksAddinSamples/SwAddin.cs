using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace HymmaSampleAddin
{
    /// <summary>
    /// Summary description for HymmaSampleAddin.
    /// </summary>
    [Guid("43f85157-13ec-476b-a57e-02fb5973e8bb"), ComVisible(true)]
    [SwAddin(
            Description = "HymmaSampleAddin description",
            Title = "HymmaSampleAddin",
            LoadAtStartup = true
            )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks solidworks = null;
        ICommandManager _commandManager = null;
        int addinCookie = 0;
        BitmapHandler iBmp;

        public const int mainCmdGroupID = 5;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;
        public const int mainItemID3 = 2;
        public const int flyoutGroupID = 91;

        #region Event Handler Variables
        /// <summary>
        /// A Hashtable of open documents as key and DocumentEventHandler as value. DocumentEventHandler is responsible to add events to a document
        ///<para><c>a hastable is a very efficient key-value pair that hashes the key first and then finds the index of the value</c></para> 
        /// </summary>
        Hashtable documentsEventsRepo = new Hashtable();
        /// <summary>
        /// this object is actually the addin itself except that it is derived from SldWorks instead of ISldWorks. So that we can access the events in the addin we 
        /// will the addin into this object and later use it to access the events.
        /// </summary>
        SolidWorks.Interop.sldworks.SldWorks addin = null;
        #endregion

        #region Property Manager Variables
        UserPMPage ppage = null;
        #endregion


        // Public Properties
        public ISldWorks SwApp
        {
            get { return solidworks; }
        }
        public ICommandManager CmdMgr
        {
            get { return _commandManager; }
        }

        public Hashtable OpenDocs
        {
            get { return documentsEventsRepo; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }
        #endregion


        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {

            solidworks = (ISldWorks)ThisSW;
            addinCookie = cookie;

            //Setup callbacks
            solidworks.SetAddinCallbackInfo(0, this, addinCookie);

            #region Setup the Command Manager
            _commandManager = solidworks.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region Setup the Event Handlers
            addin = (SolidWorks.Interop.sldworks.SldWorks)solidworks;
            documentsEventsRepo = new Hashtable();
            //this will be called only the first time the addin is loaded
            //this method will attached events to all documents that open after the addin is loaded.
            AttachSwEvents();
            //Listen for events on all currently open docs
            //we need to call this method here because sometimes user fires the addin while he has some documents open already
            //there are events that will attach event handlers to all documents but until those events are fired this call to the method will suffice
            AttachEventsToAllDocuments();
            #endregion

            #region Setup Sample Property Manager
            AddPMP();
            #endregion

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            RemovePMP();
            DetachSwEvents();
            DetachEventsFromAllDocuments();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_commandManager);
            _commandManager = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(solidworks);
            solidworks = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup commnadGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            //we will need an Assembly object to create the bitmaps for the command group items
            Assembly thisAssembly;

            //these integers are references to the command group buttons. we will use these references to add those buttons to a command tab
            int commandIndex0;
            int commnadIndex1;

            //these are strings that we will use to make command group and command tab
            //this one is the name you see for this command group once you access from Tools menu
            string titleOfCommandGroup = "This is title of Command Group";
            string ToolTip = "C# Addin";
            //this is the text you see in the command tab for example 'Features' would be the tile of Feature Tab
            string TitleOfCommandTab = "Title of Command Tab";
            //if you create a command group item in a command tab this text would be its tool-tip
            string toolTipOfCommandGroup = "Tooltip of command group";

            //we will use this array to add the command tab to each document type
            int[] documentTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};


            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());
            int commandGroupError = 0;
            bool ignorePrevious = false;

            //get the ID information stored in the registry
            object registryIDs;
            bool getDataResult = _commandManager.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[2] { mainItemID1, mainItemID2 };

            //if the IDs don't match, reset the commandGroup
            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs))
                {
                    ignorePrevious = true;
                }
            }
            //a command group is a button that once clicked on, shows a list of other commands in it. it also gets listed in the Tools menu if you want
            commnadGroup = _commandManager.CreateCommandGroup2(mainCmdGroupID, titleOfCommandGroup, toolTipOfCommandGroup, "", -1, ignorePrevious, ref commandGroupError);
            //after creating the command group you should add the bitmap photos to it. for this you should use a BitMapHandler object
            commnadGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("HymmaSampleAddin.ToolbarLarge.bmp", thisAssembly);
            commnadGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("HymmaSampleAddin.ToolbarSmall.bmp", thisAssembly);
            commnadGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("HymmaSampleAddin.MainIconLarge.bmp", thisAssembly);
            commnadGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("HymmaSampleAddin.MainIconSmall.bmp", thisAssembly);

            //Here we make the buttons in the command group. They have the callback functions names as strings
            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            commandIndex0 = commnadGroup.AddCommandItem2("CreateCube", -1, "Create a cube", "Create cube", 0, "CreateCube", "", mainItemID1, menuToolbarOption);
            //ShowPMP is the call back function here which will use UserPMPage to create a new property manager page
            commnadIndex1 = commnadGroup.AddCommandItem2("Show PMP", -1, "Display sample property manager", "Show PMP", 2, nameof(ShowPMP), nameof(EnablePMP), mainItemID2, menuToolbarOption);
            //with this you get the command group listed under the Tools menu
            commnadGroup.HasToolbar = true;
            commnadGroup.HasMenu = true;
            commnadGroup.Activate();

            // a fly-out-group is a button that when you click on it shows a list of other commands in it
            //with command groups we had to use a separate property called LargIconList and SmallIconList to assign the bitmaps to the buttons
            //here we just use those same properties and use them for this flyout group as well
            FlyoutGroup flyGroup = _commandManager.CreateFlyoutGroup(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
                commnadGroup.SmallMainIcon, commnadGroup.LargeMainIcon, commnadGroup.SmallIconList, commnadGroup.LargeIconList, "FlyoutCallback", "FlyoutEnable");
            flyGroup.AddCommandItem("FlyoutCommand 1", "hint string", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");
            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;

            bool result;
            foreach (int type in documentTypes)
            {
                CommandTab commandTab;
                commandTab = _commandManager.GetCommandTab(type, TitleOfCommandTab);

                //this code removes older tabs I had created
                try
                {
                    _commandManager.RemoveCommandTab(_commandManager.GetCommandTab(type, "New Tab"));
                    _commandManager.RemoveCommandTab(_commandManager.GetCommandTab(type, "C# Addin"));
                }
                catch { }
                //if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't match-up and the tab will be blank
                if (commandTab != null & !getDataResult | ignorePrevious)
                {
                    bool res = _commandManager.RemoveCommandTab(commandTab);
                    commandTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (commandTab == null)
                {
                    commandTab = _commandManager.AddCommandTab(type, TitleOfCommandTab);
                    CommandTabBox commandBox0 = commandTab.AddCommandTabBox();
                    //we will use these to add the command to a command box
                    // we will get the command ids from the command group items references (i.e. commnadIndex* we defined earlier)
                    int[] commandIds = new int[3];
                    int[] TextType = new int[3];

                    commandIds[0] = commnadGroup.get_CommandID(commandIndex0);

                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    commandIds[1] = commnadGroup.get_CommandID(commnadIndex1);

                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;

                    commandIds[2] = commnadGroup.ToolbarId;

                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    result = commandBox0.AddCommands(commandIds, TextType);

                    CommandTabBox commandBox1 = commandTab.AddCommandTabBox();
                    commandIds = new int[1];
                    TextType = new int[1];

                    commandIds[0] = flyGroup.CmdID;
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;
                    result = commandBox1.AddCommands(commandIds, TextType);
                    commandTab.AddSeparator(commandBox1, commandIds[0]);
                }

            }
            thisAssembly = null;
        }
        //we will call this method once we are disconnecting from solidworks
        public void RemoveCommandMgr()
        {
            iBmp.Dispose();
            _commandManager.RemoveCommandGroup(mainCmdGroupID);
            _commandManager.RemoveFlyoutGroup(flyoutGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public Boolean AddPMP()
        {
            ppage = new UserPMPage(this);
            return true;
        }

        public Boolean RemovePMP()
        {
            ppage = null;
            return true;
        }

        #endregion

        #region UI Callbacks
        public void CreateCube()
        {
            //make sure we have a part open
            string partTemplate = solidworks.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            if ((partTemplate != null) && (partTemplate != ""))
            {
                IModelDoc2 modDoc = (IModelDoc2)solidworks.NewDocument(partTemplate, (int)swDwgPaperSizes_e.swDwgPaperA2size, 0.0, 0.0);

                modDoc.InsertSketch2(true);
                modDoc.SketchRectangle(0, 0, 0, .1, .1, .1, false);
                //Extrude the sketch
                IFeatureManager featMan = modDoc.FeatureManager;
                featMan.FeatureExtrusion(true,
                        false, false,
                        (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                        0.1, 0.0,
                        false, false,
                        false, false,
                        0.0, 0.0,
                        false, false,
                        false, false,
                        true,
                        false, false);
            }
            else
            {
                solidworks.SendMsgToUser("There is no part template available. Please check your options and make sure there is a part template selected, or select a new part template.");
            }
        }


        public void ShowPMP()
        {
            if (ppage != null)
                ppage.Show();
        }

        public int EnablePMP()
        {
            if (solidworks.ActiveDoc != null)
                return 1;
            else
                return 0;
        }

        public void FlyoutCallback()
        {
            FlyoutGroup flyGroup = _commandManager.GetFlyoutGroup(flyoutGroupID);
            flyGroup.RemoveAllCommandItems();

            flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");
        }

        public int FlyoutEnable()
        {
            return 1;
        }

        public void FlyoutCommandItem1()
        {
            solidworks.SendMsgToUser("Flyout command 1");
        }

        public int FlyoutEnableCommandItem1()
        {
            return 1;
        }
        #endregion

        /// <summary>
        ///refer to this link <a href=" https://stackoverflow.com/questions/803242/understanding-events-and-event-handlers-in-c-sharp">to learn more about events</a>
        /// </summary>
        /// <returns></returns>
        //think of an event as a list of methods. you should use a delegate to add to this list, the event will not allow you to add to this list if 
        //you try to use a different delegate than what it accepts.
        //ActiveDocChangeNotify is an Event in SldWorks that accepts a delegate of type DSldWorksEvents_ActiveDocChangeNotifyEventHandler,
        //this delegate accepts methods that return int
        //https://stackoverflow.com/questions/803242/understanding-events-and-event-handlers-in-c-sharp explains it further
        #region Event Methods

        private bool AttachSwEvents()
        {
            try
            {
                addin.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                addin.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                addin.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                addin.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                addin.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        /// <summary>
        /// Get the first open document from the addin, iterates through the documents and then 
        /// if they don't have event handlers, attaches it to them based on the document type 
        /// </summary>
        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 ThisDocument = (ModelDoc2)solidworks.GetFirstDocument();
            while (ThisDocument != null)
            {
                if (!documentsEventsRepo.Contains(ThisDocument))
                {
                    AttachEventHandlersToDocument(ThisDocument);
                }
                else if (documentsEventsRepo.Contains(ThisDocument))
                {
                    DocumentEventHandler documentEventHandler = (DocumentEventHandler)documentsEventsRepo[ThisDocument];
                    if (documentEventHandler != null)
                        documentEventHandler.ConnectModelViews();
                }
                ThisDocument = (ModelDoc2)ThisDocument.GetNext();
            }
        }

        /// <summary>
        /// based on the type of modelDocument (i.e Part,Drawing,Assembly) creates a new DocumentEventHandler.cs
        /// which adds two events to the modelDocument : NewSelectionNotify and DestroyNotify
        /// method then adds the object to openDocuments hashtable (field)
        /// <para>the hashtable key is the modelDocument and the value is the DocumentEventHandler object</para>
        /// </summary>
        /// <param name="modelDocument"></param>
        /// <returns></returns>
        public bool AttachEventHandlersToDocument(ModelDoc2 modelDocument)
        {
            if (modelDocument == null) return false;

            //if modelDocument is in the hashtable, don't go through this. it means
            //this has already happened
            if (!documentsEventsRepo.Contains(modelDocument))
            {
                DocumentEventHandler documentEventHandler;
                switch (modelDocument.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            documentEventHandler = new PartEventHandler(modelDocument, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            documentEventHandler = new AssemblyEventHandler(modelDocument, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            documentEventHandler = new DrawingEventHandler(modelDocument, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                documentEventHandler.AttachEventHandlers();
                documentsEventsRepo.Add(modelDocument, documentEventHandler);
            }
            return true;
        }

        /// <summary>
        /// this detaches event handlers from addin object
        /// </summary>
        /// <returns></returns>
        private bool DetachSwEvents()
        {
            try
            {
                addin.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                addin.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                addin.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                addin.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                addin.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }

        /// <summary>
        /// this detaches event handlers from all open documents via documentsEvenHAndles
        /// </summary>
        /// <returns></returns>
        public bool DetachEventsFromAllDocuments()
        {
            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = documentsEventsRepo.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            documentsEventsRepo.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)documentsEventsRepo[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }

        /// <summary>
        /// removes the values of the <param name="modDoc">modDoc</param> from the documentsEventHandlers hashtable and sets them to null
        /// </summary>
        /// <param name="modDoc"></param>
        /// <returns></returns>
        public bool RemoveModelFromDocEventRepo(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler = (DocumentEventHandler)documentsEventsRepo[modDoc];
            documentsEventsRepo.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }
        #endregion

        #region Event Handlers
        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion
    }

}
