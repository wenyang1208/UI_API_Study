using SAPbobsCOM;
using SAPbouiCOM;
using Application = System.Windows.Forms.Application;

namespace WinFormsApp2
{
    internal static class Program
    {
        // SAPbouiCOM: Frontend operations (UI object)
        private static SAPbouiCOM.Application SBO_Application;

        private static SAPbobsCOM.Company diCompany;

        [STAThread]
        static void Main()
        {
            ConnectToUI();
            //CreateForm();
            //SaveAsXml();
            //LoadAsXml();
            SetEventFilters();
            //ExtendCancelButtonWithEvent();
            CreateMenu();
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            // Subscribe to the ItemEvent event in SAP Business One's application object
            // subscribe means that the method (SBO_Application_ItemEvent) will be executed when the item events happened
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            Application.Run();
        }

        // connection to UI API
        private static void ConnectToUI() {

            SAPbouiCOM.SboGuiApi SboGuiApi;
            string sConnectionString;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

            SboGuiApi.Connect(sConnectionString);
            SBO_Application = SboGuiApi.GetApplication();

            SBO_Application.MessageBox("Connected to UI API", 1, "Continue", "Cancel");

            // option with connection with SSO
            ConnectWithSSO();

            // option with connection with shared memory / multiple add-ons
            //ConnectWithSharedMemory();
        }

        // DI API connection via cookie
        // SSO - Single Sign On
        private static void ConnectWithSSO() {
            // 1. Get cookies from the company
            // 2. Get connection info (encrypted) from UI API with the cookie
            // 3. Convert encrypted conninfo and set login with the connection info, with ret validation

            diCompany = new SAPbobsCOM.Company();
            string cookie = diCompany.GetContextCookie(); // DI API

            // get from company in application 
            string connInfo = SBO_Application.Company.GetConnectionContext(cookie); // UI API

            int ret = diCompany.SetSboLoginContext(connInfo); // DI API

            if (ret != 0)
            {
                // In SAP B1 frontend obj
                SBO_Application.MessageBox("Connection is failed", 0, "OK", "", "");
            }
            else {
                SBO_Application.MessageBox("Connected with SSO", 0, "OK", "", "");
            }
        }

        // connect with shared memory/multiple add-on
        private static void ConnectWithSharedMemory() {
            diCompany = (SAPbobsCOM.Company) Program.SBO_Application.Company.GetDICompany();
            SBO_Application.MessageBox("Connected with shared memory " + Program.diCompany.CompanyName, 0, "OK","","");
        }
        public static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType) {
            switch (EventType) {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    SBO_Application.MessageBox("My is addon disconnected." + Program.diCompany.CompanyName, 0, "Ok", "", "");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Program.diCompany);
                    Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
        // Item Event
        // create a form
        public static void CreateForm()
        {
            SAPbouiCOM.Form oForm;
            // Specify the parameters needed to create a new Form (act like blueprint for the form)
            SAPbouiCOM.FormCreationParams creationPackage;

            // define creationPackage
            // it is parameter object that the form required
            creationPackage = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

            // add properties for the creationPackage
            creationPackage.UniqueID = "TB1_DVDAvailability";
            creationPackage.FormType = "TB1_DVDAvailability";
            // borderstyle is design-time property (must be defined before the form is created)
            creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;

            // define form object
            // pass the blueprint to the form object
            // .AddEx specify more flexible with design-time proerty (blueprint) than .Add()
            oForm = SBO_Application.Forms.AddEx(creationPackage);

            oForm.Title = "DVD Availability Check";
            oForm.Left = 400;
            oForm.Top = 100;
            oForm.ClientWidth = 270;
            oForm.ClientHeight = 154;

            // create label - DVD name
            SAPbouiCOM.Item oItem;
            SAPbouiCOM.StaticText oStatic;

            // Add(UID, Type)
            oItem = oForm.Items.Add("lb_name", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 5;
            oItem.Top = 20;
            oItem.Width = 80;
            oItem.Height = 14;
            oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific)); // Specific property of the item (which is the static text)
            oStatic.Caption = "DVD Name";

            // create label - DVD Aisle
            oItem = oForm.Items.Add("lb_aisle", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 5;
            // if two label have the same position, the latest one will overwrite the previous one
            oItem.Top = 39;
            oItem.Width = 80;
            oItem.Height = 14;
            oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStatic.Caption = "DVD Aisle";

            // create label - DVD Section
            oItem = oForm.Items.Add("lb_section", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 5;
            oItem.Top = 58;
            oItem.Width = 80;
            oItem.Height = 14;
            oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStatic.Caption = "DVD Section";


            // create label - DVD Rented
            oItem = oForm.Items.Add("lb_rented", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 5;
            oItem.Top = 77;
            oItem.Width = 80;
            oItem.Height = 14;
            oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStatic.Caption = "DVD Rented";



            // create label - Rented To
            oItem = oForm.Items.Add("lb_rentTo", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 5;
            oItem.Top = 96;
            oItem.Width = 80;
            oItem.Height = 14;
            oStatic = ((SAPbouiCOM.StaticText)(oItem.Specific));
            oStatic.Caption = "Rented To";

            // add text box - it_EDIT
            oItem = oForm.Items.Add("tx_name", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = 90;
            oItem.Top = 20;
            oItem.Width = 175;
            oItem.Height = 14;
            oItem.LinkTo = "lb_name";

            oItem = oForm.Items.Add("tx_aisle", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = 90;
            oItem.Top = 39;
            oItem.Width = 175;
            oItem.Height = 14;
            oItem.LinkTo = "lb_aisle";

            oItem = oForm.Items.Add("tx_section", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = 90;
            oItem.Top = 58;
            oItem.Width = 175;
            oItem.Height = 14;
            oItem.LinkTo = "lb_section";

            oItem = oForm.Items.Add("tx_rented", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = 90;
            oItem.Top = 77;
            oItem.Width = 175;
            oItem.Height = 14;
            oItem.LinkTo = "lb_rented";


            oItem = oForm.Items.Add("tx_rentTo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oItem.Left = 90;
            oItem.Top = 96;
            oItem.Width = 175;
            oItem.Height = 14;
            oItem.LinkTo = "lb_rentTo";

            // Add button
            SAPbouiCOM.Button oButton;
            // 1 : OK
            oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Top = 130;
            oItem.Width = 65;
            oItem.Height = 19;

            // 1 : Cancel
            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 75;
            oItem.Top = 130;
            oItem.Width = 65;
            oItem.Height = 19;

            // Rent To button
            oItem = oForm.Items.Add("rentTo", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 200;
            oItem.Top = 130;
            oItem.Width = 65;
            oItem.Height = 19;
            oButton = (SAPbouiCOM.Button)oItem.Specific;
            oButton.Caption = "Rent DVD";

            oForm.Visible = true;
        }

        // save form as XML
        public static void SaveAsXml() {
            SAPbouiCOM.Form oForm;
            oForm = SBO_Application.Forms.GetForm("TB1_DVDAvailability", 0);
            
            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
            string sXmlString = oForm.GetAsXML();

            // converting the XML representation of the SAP Business One
            // form into a structured format that the XmlDocument object can work with.
            oXmlDoc.LoadXml(sXmlString);
            oXmlDoc.Save("../../../DVDAvailabilty.xml");
        }

        // Load the form from xml
        public static void LoadAsXml() {
            try
            {
                SAPbouiCOM.Form oForm;
                System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
                SAPbouiCOM.FormCreationParams creationPackage;

                oXmlDoc.Load("../../../DVDAvailabilty.xml");
                creationPackage = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

                // InnerXml: The raw string of XML in the oXmlDoc
                creationPackage.XmlData = oXmlDoc.InnerXml;

                oForm = SBO_Application.Forms.AddEx(creationPackage);
                CreateContextMenu(oForm);
                DataBinding(oForm);
                oForm.Visible = true;
            }
            catch (Exception ex) {
                SBO_Application.MessageBox("Exception: " + ex.Message);
            }
        }

        // Event
        // Item event - CATCH FORMLOAD EVENT FOR SALES ORDER FORM
        // ref: ensure the update in the properties of pval reflects outside the function, since it's complex object
        // out: similar to ref, but does not require the variable being initialised before passing into the method, meaning that it can be initialised within the passed method
        // Bubble Event: Whether there are any events after the add-on logic
        // Item event: Events happens in the form/its object, pval refers to the event detail
        // FormUID: Identifier of the form/object that the events happening
        public static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pval, out bool BubbleEvent) {
            BubbleEvent = true;
            // "139" - Sales Order
            // validate if the event is the form load event of the sales order document that occurs after the form loading 
            if (pval.FormTypeEx == "139" & pval.BeforeAction == false & pval.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) {
                SBO_Application.MessageBox("Caught Order Formload Event");
            }

            if (pval.FormUID == "TB1_DVDAvailability" & pval.ItemUID == "rentTo" & pval.BeforeAction == false & pval.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) {
                SBO_Application.MessageBox("Caught click on Rent DVD Button");
            }
        }

        // Event filter: Helps the application listen only to specific events occurring on specific forms
        public static void SetEventFilters() {
            SAPbouiCOM.EventFilter oFilter;
            SAPbouiCOM.EventFilters oFilters;

            oFilters = new SAPbouiCOM.EventFilters();
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);

            oFilter.AddEx("139");
            oFilter.AddEx("TB1_DVDAvailability");
            SBO_Application.SetFilter(oFilters);
        }

        public static void ExtendCancelButtonWithEvent() {
            SAPbouiCOM.Form oForm;
            SAPbouiCOM.Button oButton;

            try
            {
                // Ensure the form is loaded
                oForm = SBO_Application.Forms.Item("TB1_DVDAvailability");
                SAPbouiCOM.Item oItem = oForm.Items.Item("2"); // "2" - Cancel button
                oButton = (SAPbouiCOM.Button)oItem.Specific;

                // happens before the actual action, 'Cancel', so b4 cancelling, the message box pop up
                oButton.PressedBefore += OButton_ClickBefore;
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox($"Error: {ex.Message}");
            }
        }

        // Subscribe function
        private static void OButton_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SBO_Application.MessageBox("Cancel button pressed.");
        }

        // Menu
        // Create the form into the menu - module
        public static void CreateMenu() {
            SAPbouiCOM.Menus moduleMenus;
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                // ModuleMenuId = 43520
                menuItem = SBO_Application.Menus.Item("43520");

                // Get the menu collection of SAP Business One
                moduleMenus = menuItem.SubMenus;

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams) SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "TB1_DVDStore";
                oCreationPackage.String = "DVD Store";
                oCreationPackage.Position = -1; // Ensures that the item is added to the bottom of the menu

                fatherMenuItem = moduleMenus.AddEx(oCreationPackage);

                // add a submenu item to the new pop-up item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "TB1_Avail";
                oCreationPackage.String = "DVD Availability";
                oCreationPackage.Position = -1;

                // add submenu under the father menu
                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex) {
                SBO_Application.StatusBar.SetText("Menu already exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        public static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pval, out bool BubbleEvent) {
            BubbleEvent = true;
            if (pval.MenuUID == "TB1_Avail" & pval.BeforeAction == true) {
                LoadAsXml();
            }
            if (pval.MenuUID == "TB1_Remove" & pval.BeforeAction == true) {
                SBO_Application.Menus.RemoveEx("TB1_DVDStore");
                SBO_Application.Menus.RemoveEx("TB1_Avail");
                SBO_Application.Menus.RemoveEx("TB_Remove");
            }
        }

        // create context menu in the DVD_Availability form
        private static void CreateContextMenu(SAPbouiCOM.Form oForm) {
            SAPbouiCOM.MenuCreationParams oCreationPackage;
            oCreationPackage = (SAPbouiCOM.MenuCreationParams) SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
            oCreationPackage.UniqueID = "TB1_Remove";
            oCreationPackage.Type = BoMenuType.mt_STRING;
            oCreationPackage.String = "Remove Menu";
            // menu in the form: context menu
            oForm.Menu.AddEx(oCreationPackage);
        }

        // data binding
        // link data source to the form
        private static void DataBinding(SAPbouiCOM.Form oForm) {

            // use DB data source

            SAPbouiCOM.DBDataSource oDBDataSource;

            oDBDataSource = oForm.DataSources.DBDataSources.Add("@TB1_VIDS");

            // rented and rent to use user data source
            oForm.DataSources.UserDataSources.Add("ds_Rented", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            oForm.DataSources.UserDataSources.Add("ds_RentTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);

            SAPbouiCOM.EditText oEditText;
            SAPbouiCOM.Item oItem;

            // null: No specific conditions or filters are applied (i.e., it retrieves all rows).
            oDBDataSource.Query(null);

            oItem = oForm.Items.Item("tx_name");
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@TB1_VIDS", "Name");

            // 3rd - column in the table
            oItem = oForm.Items.Item("tx_aisle");
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@TB1_VIDS", "U_Aisle");

            oItem = oForm.Items.Item("tx_section");
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "@TB1_VIDS", "U_Section");

            // the 2nd argument is empty because it is not linked to the database table (only a UserDataSource is used)
            oItem = oForm.Items.Item("tx_rentTo");
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "", "ds_RentTo");

            oItem = oForm.Items.Item("tx_rented");
            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
            oEditText.DataBind.SetBound(true, "", "ds_Rented");
            var rentedValue = oDBDataSource.GetValue("U_Rented", 0);
            oForm.DataSources.UserDataSources.Item("ds_Rented").ValueEx = rentedValue;
        }

        // for rented field in the business partner master data
        private static void CreateChooseFromList(SAPbouiCOM.Form oForm) { 
            // collection of choose from lists in a form
            SAPbouiCOM.ChooseFromListCollection oCFLs;
            oCFLs = oForm.ChooseFromLists;

            // one choose from list
            SAPbouiCOM.ChooseFromList oCFL;

            // creation params for choose from list
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
            oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams) SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
            oCFLCreationParams.ObjectType = "2"; // Business Partner Master
            oCFLCreationParams.UniqueID = "CardCodeCFL";

            // add blueprint into the choose from list
            oCFL = oCFLs.Add(oCFLCreationParams);

            SAPbouiCOM.Item oItem;
            SAPbouiCOM.EditText oEditText;

            oItem = oForm.Items.Item("tx_RentTo");
            oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
            oEditText.ChooseFromListUID = "CardCodeCFL";
        }
    }
}