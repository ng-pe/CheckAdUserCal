using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// -------------------------------
// ole db
using System.Data;

using System.Management;
using System.Diagnostics;
using System.Collections;
// Active directory

using System.Globalization;
using System.Threading;
using System.DirectoryServices;
using System.Security.Permissions;
using System.DirectoryServices.ActiveDirectory;

using System.Threading;
using System.Threading.Tasks;

namespace TSCal_Monitoring
{
    class TerminalServerUserCal
    {
        #region Members
        private string _LdapUri = null;
        private bool _Debug = false;

        private DataTable _dtUserCal;
        private DataTable _dtSrvCal;
        private Hashtable _htCalSrv;
        private Hashtable _htCalSrvLicCount;
        private Hashtable _htCalSrvLicInUseCount;
        private Hashtable _htCalSrvLicRegCount;
        private int _CalUsed = 0;
        private int _CalExpired = 0;
        private int _CalRegistred = 0;
        // Utilisation des crédentials de la session = false utilisation des credentials definit = true
        private bool _SrvCustCred = false;
        private string _SrvLogin = "";
        private string _SrvPaswd = "";
        private string _lastError = "";


        private ArrayList _CalSrvList;

        #endregion

        #region Properties
        public string lastError
        {
            get { return this._lastError; }
        }



        public bool debug
        {
            get { return this._Debug; }
            set { this._Debug = value; }
        }

        public ArrayList GetCalSrv
        {
            get
            {
                return _CalSrvList;
            }
        }


        public DataTable GetDtUserCal
        {
            get
            {
                return this._dtUserCal;
            }
        }
        public DataTable GetDtSrvCal
        {
            get
            {
                return this._dtSrvCal;
            }
        }



        public string LdapUri
        {
            get { return this._LdapUri; }
            set
            {
                this._LdapUri = value;
            }
        }

        public int CalUsedInAD
        {
            get { return this._CalUsed; }
        }
        public int CalExpiredInAD
        {
            get { return this._CalExpired; }
        }
        public int CalSrvReg
        {
            get { return this._CalRegistred; }
        }


        public bool CustCred
        {
            get { return this._SrvCustCred; }
        }
        #endregion


        #region public methode

        public void setLogin(string login)
        {
            this._SrvLogin = login;
            if (this._SrvPaswd != "" && this._SrvLogin != "")
            {
                this._SrvCustCred = true;
            }
        }

        public void setPassword(string password)
        {
            this._SrvPaswd = password;
            if (this._SrvPaswd != "" && this._SrvLogin != "")
            {
                this._SrvCustCred = true;
            }

        }

        #endregion


        #region WMI_Methode_Private
        private ManagementScope CreateNewManagementScope(string server)
        {
            string serverString = @"\\" + server + @"\root\cimv2";

            ManagementScope scope = new ManagementScope(serverString);

            if (this._SrvCustCred == true)
            {
                ConnectionOptions options = new ConnectionOptions
                {
                    Username = this._SrvLogin,
                    Password = this._SrvPaswd,
                    Impersonation = ImpersonationLevel.Impersonate,
                    Authentication = AuthenticationLevel.PacketPrivacy
                };
                scope.Options = options;
            }

            return scope;
        }
        #endregion

        #region Methode_Private_CalSrv
        private bool CalLicSrv_Check(string server)
        {
            // controle si le serveur à un Role TSLicenseServer est disponible
            ManagementScope scope = CreateNewManagementScope(server);
            SelectQuery query = new SelectQuery("select * from Win32_TSLicenseServer");

            try
            {
                using (var searcher = new ManagementObjectSearcher(scope, query))
                {
                    ManagementObjectCollection wmiTSLicenseServer = searcher.Get();
                    if (wmiTSLicenseServer.Count <= 1)
                    {
                        string Description = "";
                        foreach (ManagementObject queryObj in searcher.Get())
                        {
                            Description = queryObj["ProductId"].ToString();
                        }


                        this._htCalSrv[server] = Description;
                        this._htCalSrvLicCount[server] = 0;
                        this._htCalSrvLicInUseCount[server] = 0;
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                }
            }
            catch (Exception exception)
            {
                //dev-debug
                this._lastError = "ERROR " + exception.Message.ToString();
                return false;

            }

        }

        public int GetLSRegCount(string server)
        {

            try
            {

                return Convert.ToInt32(this._htCalSrvLicRegCount[server]);
            }
            catch
            {
                return -1;
            }
        }

        public int GetLSInUseCount(string server)
        {
            try
            {

                return Convert.ToInt32(this._htCalSrvLicInUseCount[server]);
            }
            catch
            {
                return -1;
            }

        }
        public int GetLSCount(string server)
        {
            try
            {

                return Convert.ToInt32(this._htCalSrvLicCount[server]);
            }
            catch
            {
                return -1;
            }

        }

        public bool CalLicSrv_GetData()
        {

            if (this._CalSrvList.Count != 0)
            {
                this.initDtSrvCal();
                this._CalRegistred = 0;

                foreach (string server in this._CalSrvList)
                {

                    // get Server ProductId from hashtable


                    string strWmiSql = @"select * from Win32_TSLicenseKeyPack where ProductType = 1 and ProductVersionID = 2";
                    ManagementScope scope = CreateNewManagementScope(server);
                    SelectQuery query = new SelectQuery(strWmiSql);
                    try
                    {
                        using (var searcher = new ManagementObjectSearcher(scope, query))
                        {
                            foreach (ManagementObject queryObj in searcher.Get())
                            {

                                //Console.WriteLine("AvailableLicenses: {0}", queryObj["AvailableLicenses"]);
                                int AvailableLicenses = Convert.ToInt32(queryObj["AvailableLicenses"].ToString());
                                this._CalRegistred += AvailableLicenses;
                                this._htCalSrvLicRegCount[server] = Convert.ToInt32(this._htCalSrvLicRegCount[server]) + AvailableLicenses;

                                //Console.WriteLine("Description: {0}", queryObj["Description"]);
                                string Description = queryObj["Description"].ToString();

                                //Console.WriteLine("ExpirationDate: {0}", queryObj["ExpirationDate"]);
                                string ExpirationDate = queryObj["ExpirationDate"].ToString();

                                // Console.WriteLine("IssuedLicenses: {0}", queryObj["IssuedLicenses"]);
                                int IssuedLicenses = Convert.ToInt32(queryObj["IssuedLicenses"].ToString());

                                //Console.WriteLine("KeyPackId: {0}", queryObj["KeyPackId"]);
                                int KeyPackId = Convert.ToInt32(queryObj["KeyPackId"].ToString());

                                //Console.WriteLine("KeyPackType: {0}", queryObj["KeyPackType"]);
                                int KeyPackType = Convert.ToInt32(queryObj["KeyPackType"].ToString());

                                //Console.WriteLine("ProductType: {0}", queryObj["ProductType"]);
                                int ProductType = Convert.ToInt32(queryObj["ProductType"].ToString());

                                //Console.WriteLine("ProductVersion: {0}", queryObj["ProductVersion"]);
                                string ProductVersion = queryObj["ProductVersion"].ToString();

                                //Console.WriteLine("ProductVersionID: {0}", queryObj["ProductVersionID"]);
                                string ProductVersionID = queryObj["ProductVersionID"].ToString();

                                //Console.WriteLine("TotalLicenses: {0}", queryObj["TotalLicenses"]);
                                int TotalLicenses = Convert.ToInt32(queryObj["TotalLicenses"].ToString());

                                string ProductId = "";
                                if (this._htCalSrv[server] != null)
                                {
                                    ProductId = this._htCalSrv[server].ToString();
                                }

                                object[] row0 = { server, ProductId, AvailableLicenses, Description, ExpirationDate, IssuedLicenses, KeyPackId, KeyPackType, ProductType, ProductVersion, ProductVersionID, TotalLicenses };

                                this._dtSrvCal.Rows.Add(row0);


                            }
                        }
                    }
                    catch (ManagementException e)
                    {
                        this._lastError = "An error occurred while querying for WMI data: " + e.Message.ToString();
                        return false;
                    }



                }
            }
            else
            {
                this._lastError = "Enable to get data from license server";
                return false;
            }

            return true;

        }


        #endregion



        public TerminalServerUserCal(bool reporting = false)
        {
            this.Initializer(reporting);

        }
        private void Initializer(bool reporting)
        {
            this._dtUserCal = new DataTable();
            this._dtSrvCal = new DataTable();
            this._CalSrvList = new ArrayList();
            this._htCalSrv = new Hashtable();
            this._htCalSrvLicCount = new Hashtable();
            this._htCalSrvLicInUseCount = new Hashtable();
            this._htCalSrvLicInUseCount["!UNKNOWN"] = 0;

            this._htCalSrvLicRegCount = new Hashtable();

        }

        public bool AddCalSrv(string calsrv)
        {
            if (CalLicSrv_Check(calsrv) == true)
            {
                this._CalSrvList.Add(calsrv);
                return true;
            }
            else
            {
                return false;
            }


        }

        private void initDtSrvCal()
        {
            this._dtSrvCal = new DataTable();
            this._dtSrvCal.Columns.Add("Server", typeof(string));
            this._dtSrvCal.Columns.Add("ProductId", typeof(string));
            this._dtSrvCal.Columns.Add("AvailableLicenses", typeof(int));
            this._dtSrvCal.Columns.Add("Description", typeof(string));
            this._dtSrvCal.Columns.Add("ExpirationDate", typeof(string));
            this._dtSrvCal.Columns.Add("IssuedLicenses", typeof(int));
            this._dtSrvCal.Columns.Add("KeyPackId", typeof(int));
            this._dtSrvCal.Columns.Add("KeyPackType", typeof(int));
            this._dtSrvCal.Columns.Add("ProductType", typeof(int));
            this._dtSrvCal.Columns.Add("ProductVersion", typeof(string));
            this._dtSrvCal.Columns.Add("ProductVersionID", typeof(string));
            this._dtSrvCal.Columns.Add("TotalLicenses", typeof(int));
        }

        private void initDtUserCal()
        {

            this._dtUserCal = new DataTable();
            this._dtUserCal.Columns.Add("TSLicenseServer", typeof(string));
            this._dtUserCal.Columns.Add("SAMAccountName", typeof(string));
            this._dtUserCal.Columns.Add("varMsTSLicenseVersion", typeof(string));
            this._dtUserCal.Columns.Add("msTSExpireDate", typeof(string));
            this._dtUserCal.Columns.Add("msTSManagingLS", typeof(string));
            this._dtUserCal.Columns.Add("Status", typeof(string));
        }


        public bool SetAutoLdapUri()
        {
            try
            {

                // Construction automatique du champ LDAP
                string current_domaine = System.Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                current_domaine.Split('.');

                string[] current_domaine_split = current_domaine.Split('.');
                int i = 1;
                string final_uri = "LDAP://";
                foreach (string element in current_domaine_split)
                {
                    string v1 = "DC=" + element;
                    if (current_domaine_split.Count() <= i)
                    {
                        // dernier element
                        final_uri = final_uri + "," + v1;
                    }
                    else if (1 == i)
                    {
                        final_uri = final_uri + v1;
                    }
                    else
                    {
                        // element internedait
                        final_uri = final_uri + "," + v1;
                    }
                    i++;

                }
                this._LdapUri = final_uri;
                return true;

            }
            catch
            {
                return false;
            }

        }

        // récupération des informations de CAL dans l'annuaire active directory.
        private string GetSystemDomain()
        {
            try
            {
                return Domain.GetComputerDomain().ToString().ToLower();
            }
            catch (Exception e)
            {
                e.Message.ToString();
                return string.Empty;
            }
        }




        public bool adresquest(bool reporting = false)
        {
            // init datatable for reporting
            if (reporting == true)
            {
                this.initDtUserCal();
            }
            // Function variable :

            int LIC_EXPIRED = 0;
            int LIC_VALIDED = 0;
            // current datetime for expiration control
            System.DateTime DateControle = DateTime.Now;

            try
            {

                DirectorySearcher dirSearch = null;

                try
                {
                    // authentification 
                    if (this._SrvCustCred == true)
                    {
                        // if credential is provided
                        dirSearch = new DirectorySearcher(new DirectoryEntry(LdapUri, this._SrvLogin, this._SrvPaswd));
                    }
                    else
                    {
                        // use current user credential
                        dirSearch = new DirectorySearcher(new DirectoryEntry(LdapUri));
                    }

                }
                catch (DirectoryServicesCOMException e)
                {
                    Console.WriteLine("Connection Creditial is Wrong!!!, please Check.");
                    Console.WriteLine(e.Message.ToString());
                    return false;
                }

                // Init search object
                dirSearch.Filter = "(&(objectCategory=person)(objectClass=user)(((msTSManagingLS=*)(msTSLicenseVersion=*)(msTSExpireDate=*))))";
                dirSearch.SearchScope = SearchScope.Subtree;
                dirSearch.ServerTimeLimit = TimeSpan.FromSeconds(90);
                dirSearch.PageSize = 1000;

                dirSearch.PropertiesToLoad.Clear();
                dirSearch.PropertiesToLoad.Add("objectGuid");

                SearchResultCollection userObjectAll = dirSearch.FindAll();



                SearchResult[] searchResults = userObjectAll.Cast<System.DirectoryServices.SearchResult>().ToArray();


                int CurrentPageSize = 1000;


                // grab users informations 
                for (var step = 0; step < Math.Ceiling((double)userObjectAll.Count / CurrentPageSize); step++)
                {
                    //Debug// Console.WriteLine("step {0},", step.ToString());
                    Parallel.ForEach(searchResults.Skip(step * CurrentPageSize).Take(CurrentPageSize), result =>
                    {
                        using (var entry = result.GetDirectoryEntry())
                        {
                            entry.RefreshCache(new[] { "SAMAccountName" });
                            if (entry.Properties.Contains("SAMAccountName"))
                            {

                                var varLogin = entry.Properties["SAMAccountName"].Value;
                                var varTerminalServer = entry.Properties["terminalServer"].Value;
                                var varMsTSManagingLS = entry.Properties["msTSManagingLS"].Value;
                                var varMsTSLicenseVersion = entry.Properties["msTSLicenseVersion"].Value;
                                var varMsTSExpireDate = entry.Properties["msTSExpireDate"].Value;

                                int IsValidPUCAL = 0;
                                string varStatus = "UNKNOWN";

                                // Check TS User CAL:
                                if (varTerminalServer == null && varMsTSManagingLS != null && varMsTSLicenseVersion != null && varMsTSExpireDate != null)
                                {


                                    // parsing date
                                    if (varMsTSExpireDate != null)
                                    {
                                        DateTime DateTimeMsTSExpireDate;
                                        if (DateTime.TryParse(varMsTSExpireDate.ToString(), out DateTimeMsTSExpireDate))
                                        {
                                            if (DateTimeMsTSExpireDate < DateControle)
                                            {
                                                varStatus = "EXPIRE";
                                                IsValidPUCAL = 2;
                                                LIC_EXPIRED++;
                                            }
                                            else
                                            {
                                                varStatus = "USED";
                                                IsValidPUCAL = 1;
                                                LIC_VALIDED++;
                                            }
                                        }
                                        else
                                        {
                                            IsValidPUCAL = 1;
                                            LIC_VALIDED++;
                                        }
                                    }
                                    else
                                    {
                                        IsValidPUCAL = 1;
                                        LIC_VALIDED++;
                                    }

                                }
                                else
                                {
                                    // This means User does not have License
                                    IsValidPUCAL = 0;
                                }




                                // ajout de la ligne dans la dataTable

                                if (varTerminalServer == null) varTerminalServer = "NC";
                                if (varMsTSManagingLS == null) varMsTSManagingLS = "NC";
                                if (varMsTSLicenseVersion == null) varMsTSLicenseVersion = "NC";
                                if (varMsTSExpireDate == null) varMsTSExpireDate = "NC";

                                string varServer = "!UNKNOWN";
                                bool FindProductId = false;

                                // ProductID search
                                foreach (DictionaryEntry dictionaryEntry in this._htCalSrv)
                                {
                                    if (dictionaryEntry.Value.ToString() == varMsTSManagingLS.ToString() && FindProductId == false)
                                    {
                                        varServer = dictionaryEntry.Key.ToString();
                                        this._htCalSrvLicCount[varServer] = Convert.ToInt32(this._htCalSrvLicCount[varServer]) + 1;
                                        if (varStatus == "USED")
                                        {
                                            this._htCalSrvLicInUseCount[varServer] = Convert.ToInt32(this._htCalSrvLicInUseCount[varServer]) + 1;
                                        }

                                        FindProductId = true;
                                        break;
                                    }


                                }

                                if (varStatus == "USED" && varServer == "!UNKNOWN")
                                {
                                    this._htCalSrvLicInUseCount["!UNKNOWN"] = Convert.ToInt32(this._htCalSrvLicInUseCount["!UNKNOWN"]) + 1;
                                }

                                // 
                                if (reporting == true)
                                {
                                    string[] row0 = { varServer, varLogin.ToString(), varMsTSLicenseVersion.ToString(), varMsTSExpireDate.ToString(), varMsTSManagingLS.ToString(), varStatus.ToString() };
                                    this._dtUserCal.Rows.Add(row0);
                                    Console.WriteLine("# " + varLogin.ToString() + " - " + varStatus.ToString());
                                }


                            }
                        }


                    });
                }


                // update Licences Counter
                this._CalUsed = LIC_VALIDED;
                this._CalExpired = LIC_EXPIRED;

            }
            catch (ManagementException e)
            {
                Console.WriteLine("An error occurred : " + e.Message);
                return false;
            }

            return true;


        }





    }
}
