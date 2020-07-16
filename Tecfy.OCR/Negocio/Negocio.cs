using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using Tecfy.OCR.com.softexpert;
using iTextSharp;
using iTextSharp.text;
using System.Security.Cryptography;
using System.Net.Mail;
using System.Linq;


namespace Tecfy.OCR
{
    partial class OCR
    {
        #region .: Variables :.
        private static string BackupArquivos = ConfigurationManager.AppSettings["BackupArquivos"].ToString();
        private static string pathInput = Ready.AppSettings["pastaEntrada"].ToString();
        private static string logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
        private static string PastaArquivos = ConfigurationManager.AppSettings["PastaArquivos"].ToString();
        private static string SemPastaAluno = ConfigurationManager.AppSettings["SemPastaAluno"].ToString();
        private static string separator = ConfigurationManager.AppSettings["Separator"];
        private static string IntervalReturn = ConfigurationManager.AppSettings["IntervalReturnProcessException"];


        private static FileSystemWatcher watcher;
        private static List<string> runningFiles = new List<string>();
        private static EventLog EventLog = null;
        private static System.Timers.Timer timer = new System.Timers.Timer();
        private static DateTime lastExecution;
        private static int iCriar = 1;
        private static int iImport = 1;
        private static int iAssoc = 1;
        private static int Pesq2 = 1;

        #endregion

        #region .: Constructor :.

        public OCR()
        {
        }

        internal static void SetEventLog(EventLog eventLog)
        {
            EventLog = eventLog;
        }

        #endregion

        #region .: Start Process :.

        public static void Start()
        {
            try
            {
                CreateFolder();
                int processes = Convert.ToInt32(Ready.AppSettings["Processes"]);
                ThreadPool.GetMaxThreads(out int maxWorker, out int maxIOC);
                ThreadPool.SetMaxThreads(processes, maxIOC);

                ProcessCurrentFiles();

                InitFileSystemWatcher(1);

                timer.Elapsed += new System.Timers.ElapsedEventHandler(Restart);
                timer.Interval = Convert.ToInt32(Ready.AppSettings["Interval.Restart"]);
                timer.Enabled = true;
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(string.Format("Método Start, Erro: {0}", ex.Message), EventLogEntryType.Error);
            }
        }

        public static void Restart(object source, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                if (Directory.GetFiles(pathInput).Length > 0)
                {
                    if (lastExecution.AddMilliseconds(240000) < DateTime.Now)
                    {
                        watcher.EnableRaisingEvents = false;
                        watcher.Dispose();
                        watcher = null;
                        ProcessCurrentFiles();
                        InitFileSystemWatcher(1);
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(string.Format("Método Restart, Erro: {0}", ex.Message), EventLogEntryType.Error);
            }
        }

        private static void Run(object state)
        {
            var item = (string)state;

            string nome = Path.GetFileName(item);

            var fileNameArray = Path.GetFileName(item.Replace(".pdf", "")).ToString().Split(new char[] { Convert.ToChar(separator) });

            var GeraBkp = ConfigurationManager.AppSettings["GeraBkp"].ToString();

            var P0 = ConfigurationManager.AppSettings["IDENTIFICADOR_DOC"];
            var P1 = ConfigurationManager.AppSettings["IDCATEGORY"];
            var P2 = ConfigurationManager.AppSettings["TITULO"];
            ///var P3 = ConfigurationManager.AppSettings["CPF"];

            var P4 = ConfigurationManager.AppSettings["P_Atributo1"];
            var P5 = ConfigurationManager.AppSettings["P_Atributo2"];
            var P6 = ConfigurationManager.AppSettings["P_Atributo3"];
            var P7 = ConfigurationManager.AppSettings["P_Atributo4"];
            var P8 = ConfigurationManager.AppSettings["P_Atributo5"];
            var P9 = ConfigurationManager.AppSettings["P_Atributo6"];
            var P10 = ConfigurationManager.AppSettings["P_Atributo7"];
            var P11 = ConfigurationManager.AppSettings["P_Atributo8"];
            var P12 = ConfigurationManager.AppSettings["P_Atributo9"];
            var P13 = ConfigurationManager.AppSettings["P_Atributo10"];
            var P14 = ConfigurationManager.AppSettings["IDPASTADESTINO"];
            var P15 = ConfigurationManager.AppSettings["IDUSER"];

            try
            {
                string Identificador = "";
                if (P0 != "") { Identificador = fileNameArray[Convert.ToInt32(P0)]; }

                string IDCATEGORY = "";
                if (P1 != "") { IDCATEGORY = fileNameArray[Convert.ToInt32(P1)]; }

                string Titulo = "";
                if (P2 != "") { Titulo = fileNameArray[Convert.ToInt32(P2)]; }

                //string CPF = "";
                //if (P3 != "") { CPF = fileNameArray[Convert.ToInt32(P3)]; }

                string PosAtr1 = "";
                if (P4 != "") { PosAtr1 = fileNameArray[Convert.ToInt32(P4)]; }

                string PosAtr2 = "";
                if (P5 != "") { PosAtr2 = fileNameArray[Convert.ToInt32(P5)]; }

                string PosAtr3 = "";
                if (P6 != "") { PosAtr3 = fileNameArray[Convert.ToInt32(P6)]; }

                string PosAtr4 = "";
                if (P7 != "") { PosAtr4 = fileNameArray[Convert.ToInt32(P7)]; }

                string PosAtr5 = "";
                if (P8 != "") { PosAtr5 = fileNameArray[Convert.ToInt32(P8)]; }

                string PosAtr6 = "";
                if (P9 != "") { PosAtr6 = fileNameArray[Convert.ToInt32(P9)]; }

                string PosAtr7 = "";
                if (P10 != "") { PosAtr7 = fileNameArray[Convert.ToInt32(P10)]; }

                string PosAtr8 = "";
                if (P11 != "") { PosAtr8 = fileNameArray[Convert.ToInt32(P11)]; }

                string PosAtr9 = "";
                if (P12 != "") { PosAtr9 = fileNameArray[Convert.ToInt32(P12)]; }

                string PosAtr10 = "";
                if (P13 != "") { PosAtr10 = fileNameArray[Convert.ToInt32(P13)]; }

                string MatriculaPasta = "";
                if (P14 != "") { MatriculaPasta = fileNameArray[Convert.ToInt32(P14)]; }

                string IDUSER = "";
                if (P15 != "") { IDUSER = fileNameArray[Convert.ToInt32(P15)]; }


                string ReturnBusc = ValidMatricExistt(MatriculaPasta);

                if (ReturnBusc == "yes")
                {
                    //
                    string resposta = BusDoc.BuscarDoc(Identificador);

                    string Retornocriar = "";

                    if (resposta == "false")
                    {
                        Retornocriar = CriaDoc.CriarDoc(Identificador, IDCATEGORY, Titulo, IDUSER, PosAtr1, PosAtr2, PosAtr3, PosAtr4, PosAtr5, PosAtr6, PosAtr7, PosAtr8, PosAtr9, PosAtr10, item);

                        if (Retornocriar != "false")
                        {
                            //FIEMT
                            //var fileNameArrayMatriculaPasta = Path.GetFileName(Retornocriar).ToString().Split(new char[] { Convert.ToChar("-") });
                            //MatriculaPasta = "";
                            //

                            if (MatriculaPasta != "")
                            {
                                CriaDocContainerAssoc.CriaDocContainerAssocia(IDCATEGORY, Retornocriar, MatriculaPasta);
                            }
                            string resImport = ImportDoc.ImportarDoc(Retornocriar, item);

                            if (resImport != "false")
                            {
                                GetDestinationFolder(item, Retornocriar);
                            }
                            File.AppendAllText(logpath + @"\" + "log.txt", "\r\n ==== ==== ==== ==== ");
                        }
                    }
                    else
                    {
                        if (MatriculaPasta != "")
                        {
                            CriaDocContainerAssoc.CriaDocContainerAssocia(IDCATEGORY, Retornocriar, MatriculaPasta);
                        }
                        string resImport = ImportDoc.ImportarDoc(Identificador, item);

                        if (resImport != "false")
                        {
                            GetDestinationFolder(item, Retornocriar);
                        }
                        File.AppendAllText(logpath + @"\" + "log.txt", "\r\n ==== ==== ==== ==== ");
                    }
                }
                else
                {
                    if (ReturnBusc == "NoExist")
                    {
                        File.AppendAllText(logpath + @"\" + "log_erro_SemPastaAluno.txt", "\r\n" + DateTime.Now + @" | " + nome + @" | " + "Ausência de Pasta do Aluno" + @";");

                        if (!Directory.Exists(SemPastaAluno))
                        {
                            Directory.CreateDirectory(SemPastaAluno);
                        }

                        if (File.Exists(SemPastaAluno + "\\" + Path.GetFileName(item)))
                        {
                            File.Delete(SemPastaAluno + "\\" + Path.GetFileName(item));
                        }
                        File.Move(item, string.Concat(SemPastaAluno, "\\", Path.GetFileName(item)));
                        EnviaEmail("Ausência de Pasta do Aluno", nome);
                    }
                    else
                    {
                        File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | " + nome + @" | " + "Erro na Conexão" + @";");

                        if (File.Exists(PastaArquivos + "\\" + Path.GetFileName(item)))
                        {
                            File.Delete(PastaArquivos + "\\" + Path.GetFileName(item));
                        }
                        File.Move(item, string.Concat(PastaArquivos, "\\", Path.GetFileName(item)));
                        EnviaEmail("Erro na Conexão", nome);

                    }

                }
            }
            catch (Exception ex)
            {

                File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | " + nome + @" | " + ex.Message.ToString() + @";");
                if (File.Exists(PastaArquivos + "\\" + Path.GetFileName(item)))
                {
                    File.Delete(PastaArquivos + "\\" + Path.GetFileName(item));
                }
                File.Move(item, string.Concat(PastaArquivos, "\\", Path.GetFileName(item)));
                EnviaEmail(ex.Message.ToString(), nome);
            }
            finally
            {
                lock (runningFiles)
                {
                    runningFiles.Remove(item);
                }
            }
        }

        #endregion

        #region .: Helper :.

        private static void CreateFolder()
        {
            var pastaEntrada = ConfigurationManager.AppSettings["pastaEntrada"].ToString();
            if (!Directory.Exists(pastaEntrada))
            {
                Directory.CreateDirectory(pastaEntrada);
            }

            var PastaDestinoLog = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
            if (!Directory.Exists(PastaDestinoLog))
            {
                Directory.CreateDirectory(PastaDestinoLog);
            }

            var PastaArquivos = ConfigurationManager.AppSettings["PastaArquivos"].ToString();
            if (!Directory.Exists(PastaArquivos))
            {
                Directory.CreateDirectory(PastaArquivos);
            }

            var GeraBkp = ConfigurationManager.AppSettings["GeraBkp"].ToString();
            if (GeraBkp != "false")
            {
                var BackupArquivos = ConfigurationManager.AppSettings["BackupArquivos"].ToString();
                if (!Directory.Exists(BackupArquivos))
                {
                    Directory.CreateDirectory(BackupArquivos);
                }

            }
        }

        private static void GetDestinationFolder(string item, string Retornocriar)
        {
            bool GeraBkp = Convert.ToBoolean(ConfigurationManager.AppSettings["GeraBkp"]);
            try
            {
                if (GeraBkp == true)
                {
                    string nome = Path.GetFileName(item);
                    string nomeItem = Retornocriar + " - " + nome;

                    if (File.Exists(BackupArquivos + "\\" + Path.GetFileName(nomeItem)))
                    {
                        File.Delete(BackupArquivos + "\\" + Path.GetFileName(nomeItem));
                    }
                    File.Move(item, string.Concat(BackupArquivos + @"\\" + Path.GetFileName(nomeItem)));
                }
                else
                {
                    File.Delete(item);
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(string.Format("Método GetDestinationFolder, Erro: {0}", ex.Message), EventLogEntryType.Error);
            }

        }


        private static void WatcherError(object sender, ErrorEventArgs e)
        {
            watcher.EnableRaisingEvents = false;
            watcher.Dispose();
            watcher = null;
            ProcessCurrentFiles();
            InitFileSystemWatcher(1);
        }

        private static void WatcherOnChanged(object source, FileSystemEventArgs e)
        {
            if (File.Exists(e.FullPath))
            {
                if (e.ChangeType == WatcherChangeTypes.Created)
                {
                    FileCreationVerification.AddFileToCreatedFileList(e.FullPath);
                }
                if (FileCreationVerification.FileCreatedIsCompletedWrited(e.FullPath))
                {
                    QueueToProcess(e.FullPath);
                }
            }
        }

        private static void ProcessCurrentFiles()
        {
            foreach (var item in Directory.GetFiles(pathInput))
            {
                QueueToProcess(item);
            }
        }

        private static void QueueToProcess(string item)
        {
            lock (runningFiles)
            {
                if (!runningFiles.Contains(item))
                {
                    lastExecution = DateTime.Now;
                    runningFiles.Add(item);
                    //ThreadPool.QueueUserWorkItem(new WaitCallback(Run), item);
                    Run(item);
                }
            }
        }

        #region BuscaMatricula
        public static string ValidMatricExistt(string IdentificadorDOC)
        {
            Boolean PesqPastaAluno = Convert.ToBoolean(ConfigurationManager.AppSettings["PesqPastaAluno"]);

            if (PesqPastaAluno == true)
            {

                try
                {
                    SEClient SeachDoc1 = Conection.GetConnection();

                    searchDocumentReturn searchDocumentReturnT1 = new searchDocumentReturn();
                    searchDocumentFilter searchDocumentFilterT1 = new searchDocumentFilter { IDDOCUMENT = IdentificadorDOC };
                    searchDocumentReturnT1 = SeachDoc1.searchDocument(searchDocumentFilterT1, "", null);
                    Pesq2 = 1;

                    if (searchDocumentReturnT1.RESULTS.Count() > 0)
                    {
                        return ("yes");
                    }
                    else
                    {
                        return ("NoExist");
                    }

                }
                catch (Exception ex)
                {
                    while (Pesq2 <= 2)
                    {
                        Pesq2++;
                        File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | ValidMatricExistt | " + @" Valida Matricula: " + IdentificadorDOC + @" - " + ex.Message.ToString() + @";");
                        Thread.Sleep(Convert.ToInt32(IntervalReturn));
                        return ValidMatricExistt(IdentificadorDOC);
                    }
                    //File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | SeachDoc " + ex.Message.ToString() + @";");
                    return ("NoConect");
                }
            }
            else
            {
                return ("yes");
            }
        }
        #endregion


        #region SEARCHDOC
        public class BusDoc
        {

            public static string BuscarDoc(string Identificador)
            {

                if (Identificador == "")
                {
                    return "false";
                }

                documentDataReturn documentDataReturn = new documentDataReturn();

                SEClient SeachDoc = Conection.GetConnection();
                documentDataReturn = SeachDoc.viewDocumentData(Identificador, "", "", "");

                if (documentDataReturn.ERROR == null)
                {
                    return "verdadeiro";
                }
                else
                {
                    return "false";
                }
            }
        }
        #endregion

        #region CREATEDOC
        public class CriaDoc
        {
            public static string CriarDoc(string Identificador, string IDCATEGORY, string Titulo, string IDUSER, string PosAtr1, string PosAtr2, string PosAtr3, string PosAtr4, string PosAtr5, string PosAtr6, string PosAtr7, string PosAtr8, string PosAtr9, string PosAtr10, string item)
            {
                string nome = Path.GetFileName(item);

                try
                {
                    //NAME_ATRIBUTE
                    //var AtributoCPF = ConfigurationManager.AppSettings["AtributoCPF"];
                    var Atributo1 = ConfigurationManager.AppSettings["Atributo1"];
                    var Atributo2 = ConfigurationManager.AppSettings["Atributo2"];
                    var Atributo3 = ConfigurationManager.AppSettings["Atributo3"];
                    var Atributo4 = ConfigurationManager.AppSettings["Atributo4"];
                    var Atributo5 = ConfigurationManager.AppSettings["Atributo5"];
                    var Atributo6 = ConfigurationManager.AppSettings["Atributo6"];
                    var Atributo7 = ConfigurationManager.AppSettings["Atributo7"];
                    var Atributo8 = ConfigurationManager.AppSettings["Atributo8"];
                    var Atributo9 = ConfigurationManager.AppSettings["Atributo9"];
                    var Atributo10 = ConfigurationManager.AppSettings["Atributo10"];
                    var DSRESUME = ConfigurationManager.AppSettings["DSRESUME"];


                    //ATRIBUTE_VALORES_FIXO
                    var V_Atributo1 = ConfigurationManager.AppSettings["V_Atributo1"];
                    var V_Atributo2 = ConfigurationManager.AppSettings["V_Atributo2"];
                    var V_Atributo3 = ConfigurationManager.AppSettings["V_Atributo3"];
                    var V_Atributo4 = ConfigurationManager.AppSettings["V_Atributo4"];
                    var V_Atributo5 = ConfigurationManager.AppSettings["V_Atributo5"];
                    var V_Atributo6 = ConfigurationManager.AppSettings["V_Atributo6"];
                    var V_Atributo7 = ConfigurationManager.AppSettings["V_Atributo7"];
                    var V_Atributo8 = ConfigurationManager.AppSettings["V_Atributo8"];
                    var V_Atributo9 = ConfigurationManager.AppSettings["V_Atributo9"];
                    var V_Atributo10 = ConfigurationManager.AppSettings["V_Atributo10"];
                    var IDUSERFIXED = ConfigurationManager.AppSettings["IDUSERFIXED"];

                    if (IDUSER == "")
                    {
                        IDUSER = IDUSERFIXED;
                    }

                    string FormatoData = ConfigurationManager.AppSettings["FormatoData"];
                    string DTDOCUMENT = Convert.ToString(System.DateTime.Now.ToString(FormatoData));

                    //string STEP = ConfigurationManager.AppSettings["STEP"];


                    int FGMODEL = 1;
                 

                    //if (Atributo1 != "")
                    //{
                    //    PosAtr1 = Convert.ToUInt64(PosAtr1).ToString(@"000\.000\.000\-00");
                    //}
                    //Atributo6 = "";

                    //string PosAtrAJust = PosAtr1;
                    //Atributo6 = "";
                    //PosAtr6 = "";

                    //participantsData[] participantsData = new participantsData[1];
                    //participantsData[0] = new participantsData
                    //{
                    //    STEP = STEP,

                    //};



                    string StrAtribut = (Atributo1 + "=" + PosAtr1 + "" + V_Atributo1 + ";"
                                       + Atributo2 + "=" + PosAtr2 + "" + V_Atributo2 + ";"
                                       + Atributo3 + "=" + PosAtr3 + "" + V_Atributo3 + ";"
                                       + Atributo4 + "=" + PosAtr4 + "" + V_Atributo4 + ";"
                                       + Atributo5 + "=" + PosAtr5 + "" + V_Atributo5 + ";"
                                       + Atributo6 + "=" + PosAtr6 + "" + V_Atributo6 + ";"
                                       + Atributo7 + "=" + PosAtr7 + "" + V_Atributo7 + ";"
                                       + Atributo8 + "=" + PosAtr8 + "" + V_Atributo8 + ";"
                                       + Atributo9 + "=" + PosAtr9 + "" + V_Atributo9 + ";"
                                       + Atributo10 + "=" + PosAtr10 + "" + V_Atributo10 + ";"
                                       );

                    //string Titulo1 = Titulo + " - " + PosAtr2 + "-" + PosAtr3;

                    string ATTRIBUTTES = StrAtribut.Replace("=;=;=;=;=", "").Replace("=;", "").Replace("=;=", "").Replace("=;=;=", "").Replace("=;=;=;=", "");

                    SEClient newDoc = Conection.GetConnection();
                    var resultadoNewDoc = newDoc.newDocument(IDCATEGORY, Identificador, Titulo, DSRESUME, DTDOCUMENT, ATTRIBUTTES, IDUSER, null, FGMODEL, null);

                    Thread.Sleep(1000);

                    if (Identificador == "")
                    {
                        Identificador = resultadoNewDoc.Substring(3).Replace(": Documento criado com sucesso", "");
                    }

                    if (resultadoNewDoc[0].ToString().Contains("0"))
                    {
                        File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | " + nome + @" | " + resultadoNewDoc.Substring(3).ToString() + @";");

                        if (File.Exists(PastaArquivos + "\\" + Path.GetFileName(item)))
                        {
                            File.Delete(PastaArquivos + "\\" + Path.GetFileName(item));
                        }
                        File.Move(item, string.Concat(PastaArquivos, "\\", Path.GetFileName(item)));

                        EnviaEmail(resultadoNewDoc.Substring(3), nome);

                        return "false";
                    }
                    else
                    {
                        File.AppendAllText(logpath + @"\" + "log.txt", "\r\n" + DateTime.Now + @" | " + nome + @" | " + resultadoNewDoc.Substring(3).ToString() + @";");
                    }
                    iCriar = 1;
                    return Identificador;

                }
                catch (Exception ex)
                {
                    while (iCriar <= 2)
                    {
                        iCriar++;
                        File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | CriarDoc " + @" | " + nome + @" | " + ex.Message.ToString() + @";");
                        Thread.Sleep(Convert.ToInt32(IntervalReturn));
                        return CriarDoc(Identificador, IDCATEGORY, Titulo, IDUSER, PosAtr1, PosAtr2, PosAtr3, PosAtr4, PosAtr5, PosAtr6, PosAtr7, PosAtr8, PosAtr9, PosAtr10, item);
                    }
                    //File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | CriarDoc " + ex.Message.ToString() + @";");

                    if (File.Exists(PastaArquivos + "\\" + Path.GetFileName(item)))
                    {
                        File.Delete(PastaArquivos + "\\" + Path.GetFileName(item));
                        EnviaEmail(ex.Message.ToString(), nome);
                    }
                    File.Move(item, string.Concat(PastaArquivos, "\\", Path.GetFileName(item)));
                    EnviaEmail(ex.Message.ToString(), nome);
                    iCriar = 1;
                    return "false";
                }

            }

        }
        #endregion

        #region IMPORTDOC
        public class ImportDoc
        {
            public static string ImportarDoc(string Identificador, string item)
            {
                string nome = Path.GetFileName(item);

                try
                {
                    byte[] fileBinary = File.ReadAllBytes(item);
                    eletronicFile[] Arquivo = new eletronicFile[1];
                    Arquivo[0] = new eletronicFile
                    {
                        BINFILE = fileBinary,
                        ERROR = "",
                        CONTAINER = "",
                        NMFILE = nome
                    };
                    #region TST
                    SEClient seClient = Conection.GetConnection();

                    if (fileBinary.Length > 20000000)
                    {
                        string IntervalTimeout = ConfigurationManager.AppSettings["IntervalTimeout"];

                        seClient.Timeout = (Convert.ToInt32(IntervalTimeout));
                    }
                    #endregion
                    string resultado = seClient.uploadEletronicFile(Identificador, "", "", Arquivo);
                    Thread.Sleep(1000);

                    if (resultado[0].ToString().Contains("0"))
                    {
                        File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | " + nome + @" | " + resultado.Substring(3).ToString() + @";");

                        if (File.Exists(PastaArquivos + "\\" + Path.GetFileName(item)))
                        {
                            File.Delete(PastaArquivos + "\\" + Path.GetFileName(item));
                        }
                        File.Move(item, string.Concat(PastaArquivos, "\\", Path.GetFileName(item)));

                        EnviaEmail(resultado.Substring(3).ToString(), nome);
                        return "false";
                    }
                    else
                    {
                        File.AppendAllText(logpath + @"\" + "log.txt", "\r\n" + DateTime.Now + @" | " + nome + @" | Importação: " + Identificador + @" - " + resultado.Substring(3).ToString() + @";");
                    }
                    iImport = 1;
                    return Identificador;
                }
                catch (Exception ex)
                {
                    while (iImport <= 2)
                    {
                        iImport++;
                        File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | ImportarDoc " + @" | " + nome + @" | " + ex.Message.ToString() + @";");
                        Thread.Sleep(Convert.ToInt32(IntervalReturn));
                        return ImportarDoc(Identificador, item);
                    }

                    //File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | ImportarDoc " + ex.Message.ToString() + @";");

                    if (File.Exists(PastaArquivos + "\\" + Path.GetFileName(item)))
                    {
                        File.Delete(PastaArquivos + "\\" + Path.GetFileName(item));
                    }
                    File.Move(item, string.Concat(PastaArquivos, "\\", Path.GetFileName(item)));
                    EnviaEmail(ex.Message.ToString(), nome);
                    iImport = 1;
                    return "false";
                }
            }
        }
        #endregion

        #region CREATEDOCCONTAINER
        public class CriaDocContainerAssoc
        {
            public static string CriaDocContainerAssocia(string IDCATEGORY, string Retornocriar, string MatriculaPasta)
            {
                try
                {
                    string UpperLevelCategoryID = ConfigurationManager.AppSettings["CATEGORIADEASSOCIACAO"];
                    string StructID = ConfigurationManager.AppSettings["CONTAINER"];

                    SEClient newDocContainerAss = Conection.GetConnection();
                    string resultadoNewDoc = newDocContainerAss.newDocumentContainerAssociation(UpperLevelCategoryID, MatriculaPasta, "", StructID, IDCATEGORY, Retornocriar, out long codeAssociation, out string detailAssociation);

                    if (resultadoNewDoc.Contains("SUCCESS"))
                    {
                        File.AppendAllText(logpath + @"\" + "log.txt", "\r\n" + DateTime.Now + @" | Vinculado: " + Retornocriar + @" ao " + MatriculaPasta + @" " + resultadoNewDoc.ToString() + @";");
                    }
                    else
                    {
                        File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | Vinculado: " + Retornocriar + @" ao " + MatriculaPasta + @" " + resultadoNewDoc.ToString() + @";");
                    }
                    iAssoc = 1;
                    return "";
                }
                catch (Exception ex)
                {
                    while (iAssoc <= 2)
                    {
                        iAssoc++;
                        File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | AssociacaoDoc " + @" | " + Retornocriar + @" ao " + MatriculaPasta + @" " + ex.Message.ToString() + @";");
                        Thread.Sleep(Convert.ToInt32(IntervalReturn));
                        return CriaDocContainerAssocia(IDCATEGORY, Retornocriar, MatriculaPasta);
                    }
                    EnviaEmail(ex.Message.ToString(), "Erro na Associação");
                    //File.AppendAllText(logpath + @"\" + "log_erro.txt", "\r\n" + DateTime.Now + @" | AssociacaoDoc " + ex.Message.ToString() + @";");
                    iAssoc = 1;
                    return "";
                }
            }
        }
        #endregion

        #region SEClient

        public class SEClient : Documento
        {
            private string m_HeaderName;
            private string m_HeaderValue;

            protected override WebRequest GetWebRequest(Uri uri)
            {
                HttpWebRequest request = (HttpWebRequest)base.GetWebRequest(uri);

                if (null != this.m_HeaderName)
                    request.Headers.Add(this.m_HeaderName, this.m_HeaderValue);
                return (WebRequest)request;
            }

            public void SetRequestHeader(string headerName, string headerValue)
            {
                this.m_HeaderName = headerName;
                this.m_HeaderValue = headerValue;
            }

            public void SetAuthentication(string userName, string password)
            {
                string usernamePassword = userName + ":" + password;

                this.SetRequestHeader("Authorization", "Basic " + Convert.ToBase64String(new ASCIIEncoding().GetBytes(usernamePassword)));
            }

        }
        #endregion

        #region Conection
        public class Conection
        {
            readonly static string UsernameEntrada = ConfigurationManager.AppSettings["Username"];
            readonly static string PasswordEntrada = ConfigurationManager.AppSettings["Password"];



            readonly static string URL = ConfigurationManager.AppSettings["Url"];

            public static SEClient GetConnection()
            {

                string Username = Decrypt(UsernameEntrada).Replace(" ", "");
                string Password = Decrypt(PasswordEntrada).Replace(" ", "");

                SEClient seClient = new SEClient { Url = URL };
                seClient.SetAuthentication(Username, Password);

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
                ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

                return seClient;
            }
        }
        #endregion

        #region Cript
        public static string Decrypt(string cipherText)
        {
            string EncryptionKey = "MAKV2SPBNI99212";
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }


        #endregion

        #region SendMail
        public static void EnviaEmail(string documento, string msg)
        {
            Boolean TravaEMAIL = Convert.ToBoolean(ConfigurationManager.AppSettings["TravaEMAIL"]);

            if (TravaEMAIL == true)
            {
                var EmailEntrada = ConfigurationManager.AppSettings["Email"];
                var SenhaEntrada = ConfigurationManager.AppSettings["Senha"];
                var EmailDestino = ConfigurationManager.AppSettings["EmailDestino"];
                var Mensagem = ConfigurationManager.AppSettings["Mensagem"];

                string Email = Decrypt(EmailEntrada).Replace(" ", "");
                string Senha = Decrypt(SenhaEntrada).Replace(" ", "");

                int var = 1;
                try
                {
                    var client = new SmtpClient("smtp.gmail.com", 587)
                    {
                        Credentials = new NetworkCredential(Email, Senha),
                        EnableSsl = true
                    };

                    //string test = DateTime.Now + @" | " + documento;
                    string testbody = "ALERTA DE IMPORTAÇÂO  " + DateTime.Now + "\r\n \r\n MENSAGEM: " + documento + "\r\n \r\n DOCUMENTO:  " + msg + "\r\n O ARQUIVO FOI ENVIADO PARA PASTA DE ARQUIVOS REJEITADOS \r\n \r\n " + Mensagem + " \r\n \r\n Não responda esse e-mail, ele é automatico. \r\n by Tecfy2SE";

                    //Remetente,Destinatario,Assunto,enviaMensagem
                    client.Send("..:: ALERTA ::..<" + Email + ">", EmailDestino, "ALERTA IMPORTAÇÂO SE_SUITE", testbody);

                }
                catch (Exception ex)
                {
                    var logpath = ConfigurationManager.AppSettings["PastaDestinoLog"].ToString();
                    File.AppendAllText(logpath + @"\" + "log_erro.txt", ex.Message.ToString());
                }
            }
        }
        #endregion

        private static void InitFileSystemWatcher(int exec)
        {
            try
            {
                watcher = new FileSystemWatcher();
                watcher.InternalBufferSize = 65536;
                watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
                watcher.Filter = "*.pdf";
                watcher.Created += WatcherOnChanged;
                watcher.Changed += WatcherOnChanged;
                watcher.Error += WatcherError;
                watcher.Path = pathInput;
                watcher.IncludeSubdirectories = false;
                watcher.EnableRaisingEvents = true;
            }
            catch (Exception)
            {
                if (exec <= 5)
                {
                    exec++;
                    Thread.Sleep(3000);
                    InitFileSystemWatcher(exec);
                }
            }
        }

        class FileCreationVerification
        {
            private static List<string> createdFileList = new List<string>();
            public static void AddFileToCreatedFileList(string filepath)
            {
                lock (createdFileList)
                {
                    if (!createdFileList.Contains(filepath))
                        createdFileList.Add(filepath);
                }
            }

            public static bool FileCreatedIsCompletedWrited(string filepath)
            {
                lock (createdFileList)
                {
                    if (!createdFileList.Contains(filepath))
                        return false;
                    if (!IsFileReady(filepath))
                        return false;
                    createdFileList.Remove(filepath);
                    return true;
                }
            }

            private static bool IsFileReady(string filepath)
            {
                try
                {
                    using (var inputStream = File.Open(filepath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        return true;
                    }
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        #endregion
    }
}
