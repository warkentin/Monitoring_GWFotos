using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Collections;
using System.ServiceProcess;

namespace Monitoring_GWFotos
{
    class Program
    {
        static void Main(string[] args)
        {
            // SQL-Verbindung erstellen 
            SqlConnection myConnection = new SqlConnection("user id=sa; password=;server=dpn-svr-membrain\\SQLEXPRESS;" +
                                                           "Trusted_Connection=yes; database=SPOT; connection timeout=30");

            // SQL-Datenvariable erstellen
            SqlDataReader rdr_werk = null;

            bool FolderIsEmpty;
            string texting = "";
                        
            logging("------------------------------------------------------------------- ");

            try
            {
                // SQL-Datenbank öffnen
                myConnection.Open();
                
                // Werke auslesen                
                SqlCommand cmd_werk = new SqlCommand("select * from Werke", myConnection);
                rdr_werk = cmd_werk.ExecuteReader();
                Dictionary<string, string> werks = new Dictionary<string, string>();
                ArrayList werke = new ArrayList();
                while (rdr_werk.Read())
                {
                    werks.Add(rdr_werk[0].ToString().Trim(), rdr_werk[6].ToString().Trim());
                }
                rdr_werk.Close();                
                
                // Jedes Werk abfragen
                foreach (KeyValuePair<string, string> kvp in werks)
                {
                    try
                    {
                        if (kvp.Value != "")
                        {   
                            // Pürfen ob im Verzeichnis Dateien verfügbar sind
                            FolderIsEmpty = isDirectoryEmpty(kvp.Value);
                            if (FolderIsEmpty == true) texting = "Das Verzeichnis ist leer!";
                            else texting = "Es sind Fotos zur manuellen Ablage vorhanden!";

                            // Log-Datei füllen
                            logging(DateTime.Now + " - " + kvp.Key + ": " + texting);

                            if (FolderIsEmpty == false)
                            {
                                if (!System.IO.File.Exists(@"C:\Temp\GW-Monitor\alarm-" + kvp.Key + ".txt"))
                                {
                                    Eskalation(kvp.Key, kvp.Value);
                                }
                            }
                            else
                            {
                                // Prüfung, ob ein aktueller Fall vorlag
                                if (System.IO.File.Exists(@"C:\Temp\GW-Monitor\alarm-" + kvp.Key + ".txt"))
                                {
                                    DeEskalation(kvp.Key);
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        logging(e.Message);
                    }
                }
            }
            catch (Exception e)
            {
                // SQL-Datenbank lässt sich nicht öffnen
                logging("Fehler beim Öffnen der Datenbank SPOT in Diepenau\n");
                logging(e.ToString());
            }
            finally
            {
                // SQL-Verbindung trennen / schließen                
                if (myConnection != null) myConnection.Close();
            }

            // Am Ende des Tages gehört diese Datei ins Archiv
            if (DateTime.Now.Hour == 23 && DateTime.Now.Minute > 51)
            {
                System.IO.File.Copy(@"C:\Temp\GW-Monitor\log.txt", @"C:\Temp\GW-Monitor\Archiv\"
                                    + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + "_" + "log.txt", true);
                System.IO.File.Delete(@"C:\Temp\GW-Monitor\log.txt");
            }
        }

        static void DeEskalation(string Werk_Alias)
        {
            // Log-Datei füllen
            logging(DateTime.Now + " - " + Werk_Alias + ": Keine Fotos mehr vorhanden!");

            // Alarm-Datei löschen            
            System.IO.File.Delete(@"C:\Temp\GW-Monitor\alarm-" + Werk_Alias + ".txt");

            // Es wird eine Entwarnungs-Mail verschickt
            mail_versenden("GW-monitor@polipol.de", "m.warkentin@polipol.de", "GW-Fotos manuelle Archivierung - ERLEDIGT - "
                                            + Werk_Alias, "Hallo GW-Admin,\n" + "aktuell sind keine Fotos für Werk " + Werk_Alias + " zu archivieren!\n\n "
                                            + "Viele Grüße\nd.3");
        }        

        static void Eskalation(string Werk_Alias, string Werk_Pfad)
        {
            // Es wird eine Mail verschickt, zusätzlich wird eine Datei dafür angelegt
            mail_versenden("GW-monitor@polipol.de", "m.warkentin@polipol.de", "GW-Fotos manuelle Archivierung "
                                            + Werk_Alias, "Hallo GW-Admin,\n" + "bitte Fotos archivieren für Werk " + Werk_Alias + "!\n\n "
                                            + "Dateien liegen hier: " + Werk_Pfad + "\n\n"
                                            + "Viele Grüße\nd.3");

            // Log-Datei füllen
            logging(DateTime.Now + " - " + Werk_Alias + ": Alarm-E-Mail wurde verschickt!");

            // Alarm-Datei anlegen
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Temp\GW-Monitor\alarm-" + Werk_Alias + ".txt", true))
            {
                file.WriteLine(DateTime.Now + " - " + Werk_Alias.PadRight(10, ' ') + ": - E-Mail wurde verschickt!");
            }
        }

        static void logging(string message)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Temp\GW-Monitor\log.txt", true))
            {
                file.WriteLine(message);
            }
        }

        static void mail_versenden(string absender, string empfänger, string betreff, string mail_text)
        {
            SmtpClient m = new SmtpClient();
            m.Host = "192.168.48.19";
            m.Port = 25;
            m.Send(absender, empfänger, betreff, mail_text);
        }
        static bool isDirectoryEmpty(string strPath)
        {
            return !System.IO.Directory.EnumerateFileSystemEntries(strPath).Any();
        }
    }
}