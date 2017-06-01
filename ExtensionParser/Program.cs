using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExtensionParser
{

    public class Program
    {
        static string mPercorso_Log;
        public static StreamWriter w = File.AppendText(percorso_Log + "log.txt");

        public static string percorso_Log
        {
            get
            {
                return mPercorso_Log = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);

            }
        }


        static void Main(string[] args)
        {
            //il programma vero e proprio
        }

        public static int ScriviSql(string Sigla, string Decodifica)
        {
            int ok = 0;
            SqlParameter param;
            SqlCommand command = new SqlCommand();
            command.CommandText = "sp_UpdEstensioniClienti";
            command.CommandType = CommandType.StoredProcedure;
            param = new SqlParameter("@Sigla", SqlDbType.VarChar, 255);
            param.Value = Sigla;
            command.Parameters.Add(param);
            param = new SqlParameter("@Decodifica", SqlDbType.VarChar, 5);
            param.Value = Decodifica;
            command.Parameters.Add(param);
            param = new SqlParameter("ReturnCode", SqlDbType.Int);
            param.Direction = ParameterDirection.ReturnValue;
            command.Parameters.Add(param);
            try
            {
                using (SqlConnection conn = new SqlConnection())
                {
                    SqlConnectionStringBuilder strg = new SqlConnectionStringBuilder();
                    strg["Data Source"] = "IPR-SQL01\\SQLEXPRCATALOGHI";
                    strg["Initial Catalog"] = "Cataloghi";
                    strg["User ID"] = "UserCataloghi";
                    strg["Password"] = "UserCataloghi";
                    conn.ConnectionString = strg.ToString();
                    conn.Open();
                    command.ExecuteNonQuery();


                    ok = ((int)((SqlParameter)command.Parameters["ReturnCode"]).Value);
                    conn.Close();

                }
            }
            catch (Exception ex)
            {
                ok = -2;

                using (w)
                {
                    Log(ex.Message, w);

                }

            }


            return ok;

        }

        public static void Log(string logMessage, TextWriter w)
        {
            w.Write("\r\nLog Entry : ");
            w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                DateTime.Now.ToLongDateString());
            w.WriteLine("  :");
            w.WriteLine("  :{0}", logMessage);
            w.WriteLine("-------------------------------");
        }

        //---------------------------------------------------------------------------------------------------------------------------------------------
        public void xlsParser()
        {
            string txtFileEstensioni = "";
            //StreamWriter w = File.AppendText(percorso_Log + "log.txt");
            try
            {
                ExcelReader exc1 = new ExcelReader(txtFileEstensioni);
                Dictionary<string, string> coll = new Dictionary<string, string>();
                int errorCode = exc1.GetList(ref coll);
                switch (errorCode)
                {
                    case 0:
                        foreach (String decodifica in coll.Keys)
                        {
                            if (CataloghiData.InsEstensioneCliente(coll[decodifica], decodifica))
                                Log(decodifica + "\t" + coll[decodifica] + " - Aggiunto!" + "\r\n", w);
                        }
                        break;
                    case -1:
                        //MessageBox.Show("Errore durante l'importazione", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Log("Errore durante l'importazione - Case: " + errorCode.ToString(), w);
                        break;
                    case 5003:
                        //MessageBox.Show("File non trovato", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Log("File non trovato - Case: " + errorCode.ToString(), w);
                        break;
                    case 5004:
                    case 5005:
                        //MessageBox.Show("File non corretto", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Log("File non corretto - Case: " + errorCode.ToString(), w);
                        break;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Throwed Exception: " + ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("Throwed Exception: " + ex.Message, w);
            }
            //finally { Cursor.Current = currentCursor; }
        }
        //--------------------------------------------------------------------------------------------------------




    }
    public class ExcelReader
    {
        #region Members
        #region
        const int OK = 0;
        const int ERROR = -1;
        const int FILE_NOT_FOUND = 5003;
        const int FIRST_CELL_NOT_FOUND = 5004;
        const int LAST_CELL_NOT_FOUND = 5005;
        #endregion
        #region members
        #endregion
        #region local
        string m_FilePath;
        #endregion
        #endregion

        #region CTR
        /// <summary>
        /// Costruttore
        /// </summary>
        public ExcelReader(string FilePath)
        {
            // inizializza i membri
            initMembers();
            m_FilePath = FilePath;
            checkData();
        }
        #endregion

        #region Methods
        #region private
        /// <summary>
        /// Inizializza lo stato interno
        /// </summary>
        private void initMembers()
        {
            m_FilePath = string.Empty;
        }
        /// <summary>
        /// Verifica l'esistenza del file
        /// </summary>
        /// <returns></returns>
        private bool checkData()
        {
            if (!File.Exists(m_FilePath)) return false;
            return true;
        }
        #endregion
        public int GetList(ref Dictionary<string, string> list)
        {
            //Application exapp = new Application();
            var exapp = new Microsoft.Office.Interop.Excel.Application();
            if (!checkData()) return FILE_NOT_FOUND;
            try
            {
                //             string currentSheet = "Decodifica codici cataloghi"; //"Decodifica codici cataloghi"
                object m = Type.Missing;
                Workbook oBook = exapp.Workbooks.Open(m_FilePath, m, m, m, m, m, m, m, m, m, m, m, m, m, m);
                var oSheet = oBook.Sheets.get_Item(1);

                //            Range end = ws.get_Range("A1", "B3000");
                //Microsoft.Office.Interop.Excel.Range end = null;
                var xlsRange = exapp.Range["A1:B4000"];
                object valueArray = xlsRange.Value2;
                object[,] arr = valueArray as object[,];
                int l = 0;
                int f = 0;
                for (int i1 = 1; i1 < arr.GetLength(0); i1++)
                {
                    if (arr.GetValue(i1, 1) != null)
                    {
                        if (arr.GetValue(i1, 1).ToString() == "CAMPO NOTE")
                        {
                            f = i1;
                            break;
                        }
                    }
                }
                for (int i1 = arr.GetLength(0); i1 > 0; i1--)
                {
                    if (arr.GetValue(i1, 1) != null)
                    {
                        if (arr.GetValue(i1, 1).ToString() != null)
                        {
                            l = i1;
                            break;
                        }
                    }
                }

                // recupera la prima cella
                //Range f = getFirstCell(ws);
                //// recupera l'ultima cella
                //Range l = getLastCell(ws);
                //if (l == null) return FIRST_CELL_NOT_FOUND;
                //if (f == null) return LAST_CELL_NOT_FOUND;
                for (int i1 = (f + 2); i1 < (l + 1); i1++)
                {
                    if (arr.GetValue(i1, 1) != null && arr.GetValue(i1, 2) != null)
                    {
                        string k = arr.GetValue(i1, 1).ToString();
                        string v = arr.GetValue(i1, 2).ToString();
                        if (!list.ContainsKey(k)) list.Add(k, v);
                    }
                }
                return OK;
            }
            catch (Exception ex)
            {

                //MessageBox.Show("Sorgente: " + ex.Source + "\r\n\r\n" + "Messaggio: " + ex.Message + "\r\n\r\n" + "Stack Trace: " + ex.StackTrace);
                //Program.Log("Sorgente: " + ex.Source + "\r\n\r\n" + "Messaggio: " + ex.Message + "\r\n\r\n" + "Stack Trace: " + ex.StackTrace, );
                Console.WriteLine(ex.Source + " " + ex.Message);
            }
            finally
            {
                //2013.09.17 - Paride
                // Modifica per chiusura fil eprima di quit altrimenti il file rimane "appeso" ed excel lo vede come da recuperare...
                exapp.ActiveWorkbook.Close(false, Type.Missing, Type.Missing);
                exapp.Quit();
                exapp = null;
                System.Diagnostics.Process[] proc = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process item in proc)
                {
                    item.Kill();
                }
            }
            return ERROR;
        }
    }

    public static class CataloghiData
    {
        public static bool InsEstensioneCliente(String Sigla, String Decodifica)
        {
            bool ok = true;
            //IEstensioneCliente estensione = new SqlEstensioneCliente();
            //estensione.Sigla = Sigla;
            //estensione.Decodifica = Decodifica;
            //ok = (estensione.Write() == 0);
            if (Program.ScriviSql(Sigla, Decodifica) == 0)
            {
                ok = true;
            } else
            {
                ok = false;
            }
            //ok = (Program.ScriviSql(Sigla, Decodifica)) == 0);
            return ok;
        }
    }

}
#endregion
#region IEstensioneCliente
public interface IEstensioneCliente
{
    string IdEstensione { get; set; }
    string Sigla { get; set; }
    string Decodifica { get; set; }

    void Read();
    int Write();
}
/*public class SqlEstensioneCliente : EstensioneCliente
{
    #region SqlEstensioneCliente
    public SqlEstensioneCliente()
    {
    }
    #endregion

    #region Methods
    public override void Read()
    {

    }

    /// <summary>
    /// Scrive l'estensione
    /// </summary>
    /// <returns>0: record inserito; -1: record non inserito perchè già presente; -2: errore</returns>
    public override int Write()
    {
        int ok = 0;
        SqlParameter param;
        SqlCommand command = new SqlCommand();
        command.CommandText = "sp_UpdEstensioniClienti";
        command.CommandType = CommandType.StoredProcedure;
        param = new SqlParameter("@Sigla", SqlDbType.VarChar, 255);
        param.Value = Sigla;
        command.Parameters.Add(param);
        param = new SqlParameter("@Decodifica", SqlDbType.VarChar, 5);
        param.Value = Decodifica;
        command.Parameters.Add(param);
        param = new SqlParameter("ReturnCode", SqlDbType.Int);
        param.Direction = ParameterDirection.ReturnValue;
        command.Parameters.Add(param);
        try
        {
            using (SqlManager manager = new SqlManager())
            {
                manager.Open();
                manager.ExecuteNonQuery(command);
                ok = ((int)((SqlParameter)command.Parameters["ReturnCode"]).Value);
            }
        }
        catch (Exception ex)
        {
            ok = -2;
            IToolS.Base.Utilities.Logger(ex, IToolS.Base.Utilities.LogType.IntError);
        }
        return ok;
    }
    #endregion
}*/
#endregion