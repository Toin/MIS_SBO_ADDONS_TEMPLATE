//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace MIS_SBO_ADDONS_TEMPLATE
//{
//    class Main_AddOn
//    {
//    }
//}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;

static class MainAddon
{
    public static SAPbobsCOM.Company oCompany = null/* TODO Change to default(_) if this is not a reference type */;
    public static int lRetCode = 0;
    public static int lErrCode = 0;
    public static string sErrMsg = string.Empty;

    public static string CompDB;

    private static SAPbouiCOM.Application _oApp;

    public static SAPbouiCOM.Application oApp
    {
        [MethodImpl(MethodImplOptions.Synchronized)]
        get
        {
            return _oApp;
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        set
        {
            if (_oApp != null)
            {
            }

            _oApp = value;
            if (_oApp != null)
            {
            }
        }
    }

    private static SAPbouiCOM.EventFilters _oEventFilters;

    public static SAPbouiCOM.EventFilters oEventFilters
    {
        [MethodImpl(MethodImplOptions.Synchronized)]
        get
        {
            return _oEventFilters;
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        set
        {
            if (_oEventFilters != null)
            {
            }

            _oEventFilters = value;
            if (_oEventFilters != null)
            {
            }
        }
    }

    //// Constant
    //public const static string pictCFL = "CFL.bmp";
    //public const static string Inventory_MenuId = "3072";
    public const string pictCFL = "CFL.bmp";
    public const string Inventory_MenuId = "3072";

    enum MsgBoxType
    {
        WindowMsgBox = 0,
        B1MsgBox = 1,
        B1StatusBarMsg = 2
    }


    public static void Main()
    {
        // 1. Connect via UI/DI/SSO/Multiple
        Connect();
        // '2. CreateUDTs, if not exists
        // CreateUDTs()
        // '3. RegisterUDOs. if not exists
        // RegisterUDOs()
        // 4. CreateMenus.
        // CreateMenus()
        // CreateAddOnMenus()

        // 5. UpdateMenus.
        // UpdateMenus()

        // '6. SetFilter
        // SetFilters()

        // Dim oMainForm As Form = New MainForm
        // oMainForm.ShowDialog()

        System.Windows.Forms.Application.Run();
    }

    public static void Connect()
    {
        // If My.Settings.ConnectionType.Equals("DI") Then
        // ConnectViaDISample()
        // ElseIf My.Settings.ConnectionType.Equals("UI") Then
        // ConnectViaUI()
        // ElseIf My.Settings.ConnectionType.Equals("MultiAddOn") Then
        // ConnectViaMultipleAddon()
        // Else
        // ConnectViaSSO()
        // End If

        // ConnectViaDISample()

        // ConnectViaUI()

        // ConnectViaMultipleAddon()

        ConnectViaSSO();
    }

    public static void ConnectViaDISample()
    {
        // ConnectViaDI("toin-pc", "toin-pc:30000", SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008, _
        // "maruni", "sa", "P@ssw0rd", "manager", "gk88", SAPbobsCOM.BoSuppLangs.ln_English)
        ConnectViaDI("toin-pc", "toin-pc:30000", SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008, "SBODemoUS", "sa", "P@ssw0rd", "manager", "1234", SAPbobsCOM.BoSuppLangs.ln_English);
    }

    public static void ConnectViaDI(string server, string licSrv, SAPbobsCOM.BoDataServerTypes dbType, string companyDB, string dbUser, string dbPassword, string userName, string password, SAPbobsCOM.BoSuppLangs language, string addonID = "")
    {
        try
        {
            oCompany = new SAPbobsCOM.Company();
            oCompany.Server = server;
            oCompany.LicenseServer = licSrv;
            oCompany.DbServerType = dbType;
            oCompany.DbUserName = dbUser;
            oCompany.DbPassword = dbPassword;
            oCompany.CompanyDB = companyDB;
            oCompany.UserName = userName;
            oCompany.Password = password;
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;

            if (string.IsNullOrEmpty(addonID) == false)
                oCompany.AddonIdentifier = addonID;

            oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;

            lRetCode = oCompany.Connect();

            DIErrHandler("Connectiong Company");

            CompDB = oCompany.CompanyDB;
        }

        // If lRetCode <> 0 Then
        // oCompany.GetLastError()
        // End If

        catch (Exception ex)
        {
            throw ex;
        }
    }

    public static void DIErrHandler(string action)
    {
        try
        {
            string msg;

            if (lRetCode == 0)
                msg = string.Format("{0} Succeeded", action);
            else
            {
                oCompany.GetLastError(lErrCode, sErrMsg);
                msg = string.Format("{0} failed. ErrCode: {1}. ErrMsg: {2}", action, lErrCode, sErrMsg);
            }
            MsgBoxWrapper(msg);
        }
        catch (Exception ex)
        {
            MsgBoxWrapper(ex.Message);
        }
    }

    public static void MsgBoxWrapper(string msg, MsgBoxType msgboxType = MsgBoxType.B1StatusBarMsg, SAPbouiCOM.BoStatusBarMessageType msgType = SAPbouiCOM.BoStatusBarMessageType.smt_None)
    {
        if (!(oApp == null))
        {
            if (msgboxType == MainAddon.MsgBoxType.B1MsgBox)
                oApp.MessageBox(msg);
            else if (msgboxType == MainAddon.MsgBoxType.B1StatusBarMsg)
            {
                bool isErr = (msgType == SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                oApp.SetStatusBarMessage(msg, SAPbouiCOM.BoMessageTime.bmt_Medium, isErr);
            }
            else
                Interaction.MsgBox(msg);
        }
    }

    public static void ConnectViaUI()
    {
        try
        {
            SAPbouiCOM.SboGuiApi uiAPI = new SAPbouiCOM.SboGuiApi();
            string sConnStr =  Environment.GetCommandLineArgs().GetValue(1).ToString();

            uiAPI.Connect(sConnStr);

            oApp = uiAPI.GetApplication();

            // delegate the event handler
            oApp4AppEventHandler = oApp;
            oApp4ItemEvent = oApp;
            oApp4FormData = oApp;
            oApp4MenuEvent = oApp;


            oEventFilters = new SAPbouiCOM.EventFilters();

            MsgBoxWrapper("UI API Connected.", MsgBoxType.B1StatusBarMsg, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        // uiAPI = Nothing

        catch (Exception ex)
        {
            MsgBoxWrapper(ex.Message);
        }
    }

    public static void ConnectViaSSO()
    {
        try
        {
            // 1. Connect to UI
            ConnectViaUI();

            oCompany = new SAPbobsCOM.Company();
            string sCookie = oCompany.GetContextCookie();

            string connInfo = oApp.Company.GetConnectionContext(sCookie);

            // It will set Server, db, username, password to the DI Company
            oCompany.SetSboLoginContext(connInfo);

            lRetCode = oCompany.Connect();
            MsgBoxWrapper("Addon MIS_TO_PROD Connected via SSO");

            CompDB = oCompany.CompanyDB;
        }
        catch (Exception ex)
        {
            MsgBoxWrapper(ex.Message);
        }
    }

    public static void ConnectViaMultipleAddon()
    {
        try
        {
            ConnectViaUI();
            oCompany = oApp.Company.GetDICompany();
            MsgBoxWrapper("Connected via Multiple Addon");

            CompDB = oCompany.CompanyDB;
        }
        catch (Exception ex)
        {
            MsgBoxWrapper(ex.Message);
        }
    }

    public static string LoadFromXML(string FileName)
    {
        System.Xml.XmlDocument oXmlDoc;
        string sPath;

        oXmlDoc = new System.Xml.XmlDocument();

        // // load the content of the XML File

        sPath = System.Windows.Forms.Application.StartupPath;

        oXmlDoc.Load(sPath + @"\" + FileName);

        // // load the form to the SBO application in one batch
        return (oXmlDoc.InnerXml);
    }
}
