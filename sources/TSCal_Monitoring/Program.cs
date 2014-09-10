// Outil de controle de License CAL
// 20140910
// auteur Nicolas GOLLET
// Sous licence GPL v2

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


using System.Data;
using System.Management;
using System.DirectoryServices;
using System.Collections;

using NDesk.Options;


namespace TSCal_Monitoring
{
    class Program
    {


        static int verbosity = 0;

        static void ShowHelp(OptionSet p)
        {

            Console.WriteLine("TSCal_Monitoring.exe : RDS User License Cal Monitor");
            Console.WriteLine("Check RDS TS CAL license on Active Directory.");
            Console.WriteLine("Version 1.1 - 2014/09/10 - Nicolas GOLLET");
            Console.WriteLine("Usage: MSTSCalCheck [OPTIONS]+ message");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        static void Debug(string format, params object[] args)
        {
            if (verbosity > 0)
            {
                Console.Write("DEBUG> ");
                Console.WriteLine(format, args);
            }
        }

        
        static int setWarningCalFreePercent(string strArg){

            int result = 20;
            try
            {
                result = Convert.ToInt32(strArg);
                return result;

            }
            catch
            {
                // default warning = 20
                return 20;
            }
            
        }

        static void Main(string[] args)
        {
           
            bool show_help = false;
            bool printReport = false;
            string ldapUri = "";
            string user = "";
            string password = "";
            bool specCred = false;
            bool nsclientmode = true;
            int intWarningCalFreePercent = 20;


            List<string> tscalsrv = new List<string>();

            #region Command line parsing
            // Parsing de la ligne de commande
            var p = new OptionSet() {
                { "u|username=",    "Specific username for credential", v => user = v },
                { "p|password=",    "Specific password for credential", v => password = v },
                { "w|warning=",     "NSClient++ Warning level in percent (default 20)", v => intWarningCalFreePercent = setWarningCalFreePercent(v) },
                { "s|tscalsrvhost=","{tscalsrvhost} Host or IP of Terminal Server CAL Role Server (allow comma-separated eg. -s server1,server2,server3).", v => tscalsrv.Add (v) },
                { "l|ldapuri=",     "{ldapuri} Full ldap URI (eg. LDAP://DC=dm-lvs,DC=adds ) (default automatic discovery).", v => ldapUri = v },
                { "r|report",       "Print Reports", v => printReport = v != null },
                { "c|console",      "Enable console Mode (or disable Default NSClient++ mode)", v => nsclientmode = false },
                { "v",              "increase debug message verbosity",v => { if (v != null) ++verbosity; } },
                { "h|help",         "show this message and exit", v => show_help = v != null },
            };



            List<string> extra;
            try
            {
                extra = p.Parse(args);
            }
            catch (OptionException e)
            {
                Console.Write("CheckAdUserCal: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `CheckAdUserCal --help' for more information.");
                return;
            }

            if (user != "" && password != "")
            {
                // TODO
                specCred = true;

            }

            #endregion


            if (show_help)
            {
                ShowHelp(p);
                return;
            }

            if (nsclientmode == false)
            {
                Console.WriteLine("CheckAdUserCal v1.1 (console Mode)");
                Console.WriteLine("by Nicolas GOLLET");
                Console.WriteLine("----------------------------------");
                Console.WriteLine("Check Terminal Server User CAL on Active Directory");

            }


            // init TerminalServerUserCal Class:

            TerminalServerUserCal calcheck = new TerminalServerUserCal();
            Debug("specCred = " + specCred);

            // if Arguments command line is set :
            if (specCred == true)
            {
                calcheck.setLogin(user);
                calcheck.setPassword(password);
            }


            // if ldapUri not set, determine it automatically
            #region ldap automatic URI discovery
            if (ldapUri == "")
            {
                Debug("Determine auto LDAP URI");
                if (calcheck.SetAutoLdapUri() != true)
                {
                    Console.WriteLine("CRITICAL : Enable to determine LDAP URI, Please set it manualy...");
                    return;
                }
                else
                {
                    Debug("LDAPURI=" + calcheck.LdapUri);
                }



            }
            else
            {
                calcheck.LdapUri = ldapUri;
            }

            #endregion


            // add cal license server to Class
            #region Add Cal Server

            foreach (string server in tscalsrv)
            {
                if (server.Contains(",") == true)
                {
                    // on separe les server en par l'espace
                    string[] server_list = server.Split(',');
                    foreach (string srv in server_list)
                    {
                        if (calcheck.AddCalSrv(srv) == false)
                        {
                            if (nsclientmode == false)
                            {
                                Console.WriteLine("ERROR> Cannot add TSCal Server : " + srv);
                                Console.WriteLine("       " + calcheck.lastError);
                            }
                            
                        }
                        else
                        {
                            if (nsclientmode == false)
                            {
                                Console.WriteLine("INFO > Adding TSCal Server : " + srv);
                            }
                        }
                    }
                }
                else
                {
                    // mode normal on ajoute le serveur CAL
                    if (calcheck.AddCalSrv(server) == false)
                    {
                        Debug("Enable Adding CalSrv " + server);
                      

                    }
                    else
                    {
                        if (nsclientmode == false)
                        {
                            Console.WriteLine("INFO > Adding TSCal Server : " + server);
                        }
                    }
                }
            }

            #endregion

            // check Cal Server (get info)

            #region Grab Cal Srv info
            bool ResultCalServer = calcheck.CalLicSrv_GetData();
            

            if (ResultCalServer != true )
            {
                // impossible d'optenir les informations.
                if (nsclientmode == false)
                {
                    Console.WriteLine("ERROR > Unable to perform the test (CAL Server not available)");
                    Console.WriteLine("ERROR > " + calcheck.lastError);

                }
                else
                {
                    Console.WriteLine("UNKNOWN: Unable to perform the test (" + calcheck.lastError + ")");
                    Environment.Exit(3);
                }

            }
#endregion




            if (nsclientmode == false)
            {
                Console.WriteLine("INFO > Get Data from ActiveDirectory... ( " + calcheck.LdapUri + " )");
            }

            // Grab Active directory User Cal data
            bool ResultAdRequest = calcheck.adresquest();

            if ( ResultAdRequest != true)
            {
                // enable to get info !
                if (nsclientmode == false)
                {
                    Console.WriteLine("ERROR > Unable to perform the test (Binding ActiveDirectory not available)");
                }
                else
                {
                    // for NSClient++ Unknown error...
                    Console.WriteLine("UNKNOWN: Unable to perform the test (Binding ActiveDirectory not available)");
                    Environment.Exit(3);
                }

            }

           

            #region Full report (todo)
            // TODO : Create Excel/csv Report

            if (printReport)
            {
               
                DataTable DtUsers = calcheck.GetDtUserCal;

                foreach (DataRow row in DtUsers.Rows)
                {
                    Console.WriteLine("-----------------------------");
                    foreach (DataColumn column in DtUsers.Columns)
                    {
                        Console.Write(column.ColumnName + " = " );
                        Console.Write(row[column].ToString() + "\r\n");
                    }
                }

                Console.WriteLine("########################################");
                DataTable DtSrv = calcheck.GetDtSrvCal;

                foreach (DataRow row in DtSrv.Rows)
                {
                    Console.WriteLine("-----------------------------");
                    foreach (DataColumn column in DtSrv.Columns)
                    {
                        Console.Write(column.ColumnName + " = ");
                        Console.Write(row[column].ToString() + "\r\n");
                    }
                }



            }
            #endregion


            // Commande line mode report
            if (nsclientmode == false)
            {
                Console.WriteLine("+ Active Directory Terminal Server User License :");
                Console.WriteLine("\t - Cal Used      = " + calcheck.CalUsedInAD);
                Console.WriteLine("\t - Cal Expired   = " + calcheck.CalExpiredInAD);
                Console.WriteLine("\t - Cal Registred = " + calcheck.CalSrvReg);

                Console.WriteLine("");

                Console.WriteLine("+ Cal Server Allocation");
                int globalCalFree = 0;
                int globalCalUsed = 0;
                foreach (string server in calcheck.GetCalSrv)
                {
                    if (server != "!UNKNOWN")
                    {
                        Console.WriteLine("  LS Server : " + server);

                        Console.WriteLine("\t - license issued = " + calcheck.GetLSCount(server));
                        Console.WriteLine("\t - license used = " + calcheck.GetLSInUseCount(server));
                        globalCalUsed += calcheck.GetLSInUseCount(server);
                        Console.WriteLine("\t - registered license = " + calcheck.GetLSRegCount(server));
                        int CalFree = calcheck.GetLSRegCount(server) - calcheck.GetLSInUseCount(server);
                        globalCalFree += CalFree;

                        Console.WriteLine("\t - license free = " + CalFree);

                    }
                }
                int globalCalReg = globalCalUsed + globalCalFree;
                Console.WriteLine("");
                Console.WriteLine("+ Global Cal License Count = " + globalCalFree + " free / " + globalCalReg);
                Console.WriteLine("");
                Console.WriteLine("+ Delivered license by 'UNKNOWN' LS = " + calcheck.GetLSInUseCount("!UNKNOWN"));
                Console.WriteLine("");

                Console.WriteLine("\nPress any key to exit!");
                Console.ReadKey();
            }
            else // NSCLient++ mode
            {
                int globalCalFree = 0;
                int globalCalUsed = 0;
                foreach (string server in calcheck.GetCalSrv)
                {
                    if (server != "!UNKNOWN")
                    {
                        globalCalUsed += calcheck.GetLSInUseCount(server);
                        int CalFree = calcheck.GetLSRegCount(server) - calcheck.GetLSInUseCount(server);
                        globalCalFree += CalFree;
                  
                    }
                }
                int globalCalReg = globalCalUsed + globalCalFree;

               // free Cal percent :
                int percentFree = globalCalFree * 100 / globalCalReg;

               
                string perf = @"|'free'=" + globalCalFree + @";0;0 'registered'=" + globalCalReg + @";0;0";


                if (globalCalFree == 0)
                {
                    Console.WriteLine("CRITICAL: NO MORE CAL ( 0/" + globalCalReg + " )" + perf);
                    Environment.Exit(1);

                }
                else if (percentFree < intWarningCalFreePercent)
                {
                    Console.WriteLine("WARNING: TS USER CAL LOW less " + percentFree + "% free ( " + globalCalFree + " free /" + globalCalReg + " )" + perf);
                    Environment.Exit(2);
                }

                Console.WriteLine("OK: " + globalCalFree + " free / " + globalCalReg + perf);
                Environment.Exit(0);

            }
            



        }
    }
}
