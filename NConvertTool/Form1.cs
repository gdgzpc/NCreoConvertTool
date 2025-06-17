using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using EnoviaConnector.Actions;
using Newtonsoft.Json.Linq;
using NLog;
using pfcls;
//using Configuration = System.Configuration.Configuration;
//using ConfigurationManager = System.Configuration.ConfigurationManager;

namespace NCreoConvertTool
{
    public partial class MainForm : Form
    {
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private Dictionary<string, string> g_typeDerivedOutputsMap = new Dictionary<string, string>();  //不同类型需要哪些衍生文件
        private Dictionary<string, string> g_typeAttributesMap = new Dictionary<string, string>();      //不同类型需要有哪些属性   //add 20250211
        private string[] g_needAttributeNameList;

        private string g_integName = "PROE";
        private string g_userName;
        private string g_password;
        private string g_securityContext;
        private int iOpenWaitMin;
        private int OpenWaitCount = 1;
        private string g_openSWDocCount;
        private int iOpenSWDocCount = 1000;
        private int openSWDocCount = 1;
        int iRunTime = 100;
        private string g_zoomToFit = "false";
        private string g_pdfProfilePath;

        IpfcBaseSession session = null;
        private string g_appPath = @"C:\Program Files\PTC\Creo 5.0.0.0\Parametric\bin\parametric.exe";
        EnoviaConnector.Connector connector = null;

        private System.Timers.Timer backendTimer;//20250307

        public MainForm()
        {
            InitializeComponent();
            InitializeBackendTimer();
        }
        private void InitializeBackendTimer()
        {
            backendTimer = new System.Timers.Timer(60000);
            backendTimer.Elapsed += BackendTimer_Elapsed;
            backendTimer.AutoReset = true;
        }

        private void KillCADProcess()
        {
            logger.Info("start KillCADProcess....");
            Process[] processes = Process.GetProcessesByName("XTOP");
            foreach (Process process in processes)
            {
                try
                {
                    process.Kill();
                    process.WaitForExit(); // 等待进程退出
                }
                catch (Exception ex)
                {
                    //202503308
                    //throw new Exception($"终止 SolidWorks 进程（ID: {process.Id}）时出错: {ex.Message}");
                    logger.Error($"终止 Creo 进程（ID: {process.Id}）时出错: {ex.Message}");
                    KillCADProcessUseWMI();
                }
            }
        }
        private void KillCADProcessUseWMI()
        {
            logger.Info("start KillCADProcessUseWMI....");
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                        $"SELECT * FROM Win32_Process WHERE Name = 'XTOP'"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        obj.InvokeMethod("Terminate", null);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"终止 Creo 进程时出错: {ex.Message}");
            }
        }
        private void KillConvertProcess()
        {
            logger.Info("start KillConvertProcess....");
            Process[] processes = Process.GetProcessesByName("NCreoConvertTool");
            foreach (Process process in processes)
            {
                try
                {
                    process.Kill();
                    process.WaitForExit(); // 等待进程退出
                }
                catch (Exception ex)
                {
                    throw new Exception($"终止 NCreoConvertTool 进程（ID: {process.Id}）时出错: {ex.Message}");
                }
            }
        }
        public bool IsCADRunning()
        {
            Process[] processes = Process.GetProcessesByName("XTOP");
            return processes.Length > 0;
        }

        private static IpfcBaseSession StartCADApp(string appPath, int timeoutSec = 20)
        {
            try
            {
                // 1. 定义Creo启动参数
                string creoInstallPath = appPath; // -g:no_graphics -i:rpc_input"; // 替换为你的Creo安装路径
                string workingDir = @"C:\Users\Public\Documents"; // 工作目录
                string[] creoArgs = { "-g:no_graphics" }; // 启动参数（例如：无图形模式）

                string visiable = ConfigManager.GetAppSetting("CADVisiable");
                if (!visiable.Equals("true"))
                    creoInstallPath += " -g:no_graphics ";

                // 2. 启动新的Creo进程
                IpfcAsyncConnection asyncConnection;
                CCpfcAsyncConnection cAC = new CCpfcAsyncConnection();
                asyncConnection = cAC.Start(
                        creoInstallPath,  // Creo可执行文件路径
                        workingDir        // 工作目录
                    );

                Console.WriteLine("成功启动新的Creo进程！");

                // 获取当前会话对象
                IpfcBaseSession session = (IpfcBaseSession)asyncConnection.Session;

                // 打开图纸
                //session.ChangeDirectory(@"F:\temp\");
                //CCpfcModelDescriptor descModelCreate = new CCpfcModelDescriptor();
                //IpfcModelDescriptor descModel = descModelCreate.Create((int)EpfcModelType.EpfcMDL_PART, "cassprt0001.prt", null);
                //IpfcModel model = session.RetrieveModel(descModel); // 在Creo中显示模型
                //model.Display();

                // 3. 导出STP
                //string stpPath = @"F:\temp\cassprt0001.stp";
                ////IpfcExportOptions exportOptions = (IpfcExportOptions)new CMpfcExportOptions();
                ////exportOptions.Type = (int)EpfcExportType.EpfcEXPORT_STEP;
                //model.ExportIntf3D(stpPath, (int)EpfcExportType.EpfcEXPORT_STEP, null);
                //Console.WriteLine("STP导出成功: " + stpPath);

                //string igsPath = @"F:\temp\cassprt0001.igs";
                //model.ExportIntf3D(igsPath, (int)EpfcExportType.EpfcEXPORT_IGES_3D, null);

                //// 4. 导出JPG
                //string jpgPath = @"F:\temp\cassprt0001.jpg";
                //IpfcWindow window = session.get_CurrentWindow();
                //IpfcJPEGImageExportInstructions jpegInstrs;
                //CCpfcJPEGImageExportInstructions cjins = new CCpfcJPEGImageExportInstructions();
                //jpegInstrs = cjins.Create(10.0, 7.5);
                //IpfcRasterImageExportInstructions ins = (IpfcRasterImageExportInstructions)jpegInstrs;
                //ins.DotsPerInch = EpfcDotsPerInch.EpfcRASTERDPI_100;
                //ins.ImageDepth = EpfcRasterDepth.EpfcRASTERDEPTH_24;
                //window.ExportRasterImage(jpgPath, ins);

                //session.ChangeDirectory(@"F:\temp\");
                //CCpfcModelDescriptor descModelCreate = new CCpfcModelDescriptor();
                //IpfcModelDescriptor descModel = descModelCreate.Create((int)EpfcModelType.EpfcMDL_DRAWING, "cassprt0001.drw", null);
                //IpfcModel model = session.RetrieveModel(descModel); // 在Creo中显示模型
                //model.Display();

                //string pdfPath = @"F:\temp\cassprt0001.pdf";
                //IpfcPDFExportInstructions pdfExportInstructions = new CCpfcPDFExportInstructions().Create();
                //model.Export(pdfPath, (IpfcExportInstructions)pdfExportInstructions);

                //string dxfPath = @"F:\temp\cassprt0001.dxf";
                //IpfcDXFExportInstructions dxfExportInstructions = new CCpfcDXFExportInstructions().Create();
                //model.Export(dxfPath, (IpfcExportInstructions)dxfExportInstructions);

                // 5. 关闭模型
                ////model.Erase();
                ////session.get_CurrentWindow().Close();
                //session.get_CurrentWindow().Clear();

                return session;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to connect to Creo instance1: {ex.Message}");
            }
        }

        private void RestartCADApp()
        {
            logger.Info("enter RestartCADApp....");
            //检查如果存在CAD进程，先关闭，再打开新的
            if (IsCADRunning())
            {
                KillCADProcess();
            }

            try
            {
                //g_swPath = System.Configuration.ConfigurationManager.AppSettings["SolidWorksPath"];
                g_appPath = ConfigManager.GetAppSetting("CADPath");
                //logger.Info($"g_swPath={g_swPath}");
                session = StartCADApp(g_appPath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to connect to Creo instance1: {ex.Message}");
            }
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                RestartCADApp();

                string sRuntime = ConfigManager.GetAppSetting("Time");
                //logger.Info($"sRuntime={sRuntime}");
                if (sRuntime != "")
                {
                    iRunTime = int.Parse(sRuntime);
                    timer1.Interval = iRunTime;
                }

                g_userName = ConfigManager.GetAppSetting("UserName");
                g_password = ConfigManager.GetAppSetting("Password");
                g_securityContext = ConfigManager.GetAppSetting("SecurityContext");
                string g_openWaitMin = ConfigManager.GetAppSetting("OpenWaitMin");
                g_openSWDocCount = ConfigManager.GetAppSetting("OpenSWDocCount");
                g_zoomToFit = ConfigManager.GetAppSetting("ZoomToFit");
                g_pdfProfilePath = ConfigManager.GetAppSetting("PdfProfilePath");
                //password = EncryptionHelper.Decrypt(password);
                //logger.Info("tomcatClassPath=" + tomcatClassPath);
                logger.Info($"userName={g_userName}");
                iOpenWaitMin = int.Parse(g_openWaitMin);
                logger.Info($"iOpenWaitMin={iOpenWaitMin}");
                //timerWaitOpen.Interval = 60000;
                iOpenSWDocCount = int.Parse(g_openSWDocCount);

                //登录Enovia
                connector = new EnoviaConnector.Connector(g_integName);
                int ret = connector.EnoviaLogin(g_userName, g_password, "return", g_securityContext);
                if (ret < 0)
                {
                    throw new Exception($"Enovia Login failed! {ret}");
                }

                LoadPropertiesFromServer();

                timer1.Enabled = false;
                ExecuteTask();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                string errMsg = $"Exception in MainForm_Load. {ex.Message}";
                logger.Error(errMsg);
                //MessageBox.Show(errMsg, "运行错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }

        public Dictionary<string, string> GetPropertiesFromServer(Logger logger, string integName, string userName, string password, string propertiesFileName)
        {
            logger.Info($"enter GetPropertiesFromServer...propertiesFileName={propertiesFileName}");
            Dictionary<string, string> propsMap = new Dictionary<string, string>();
            try
            {
                string props = connector.GetPropertiesFromServer(propertiesFileName);
                //logger.Info("after get..."+props);
                if (props.Length > 0)
                {
                    JObject propObj = JObject.Parse(props);
                    string code = (string)propObj["code"];
                    if (code == "0")
                    {
                        JObject dataObj = (JObject)propObj["data"];
                        foreach (var property in dataObj.Properties())
                        {
                            propsMap.Add(property.Name, dataObj[property.Name].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Exception in GetPropertiesFromServer.{ex.Message}");
                throw new Exception($"Exception in GetPropertiesFromServer.{ex.Message}");
            }
            return propsMap;
        }

        private string[] GetNeedAttributeNameList(string fileType, Dictionary<string, string> typeAttributesMap)
        {
            Encoding iso = Encoding.GetEncoding("iso-8859-1");
            byte[] btArr = iso.GetBytes(typeAttributesMap[fileType]);
            string gbkText = Encoding.GetEncoding("GBK").GetString(btArr);
            logger.Info($"typeAttributesMap[fileType]={gbkText}");
            string[] needAttributeNameList = gbkText.Split(',');
            return needAttributeNameList;
        }
        private void LoadPropertiesFromServer()
        {
            logger.Info("enter LoadPropertiesFromServer...");
            try
            {
                if (g_typeDerivedOutputsMap.Count == 0)
                {
                    g_typeDerivedOutputsMap = GetPropertiesFromServer(logger, g_integName, g_userName, g_password, "cassCADDerived_PROE.properties");
                    if (g_typeDerivedOutputsMap.Count == 0)
                    {
                        throw new System.Exception("远程属性（不同类型需要哪些衍生文件）未获得！");
                    }
                }
                if (g_typeAttributesMap.Count == 0)//add 20250211
                {
                    g_typeAttributesMap = GetPropertiesFromServer(logger, g_integName, g_userName, g_password, "cassCADAttribute_PROE.properties");
                    if (g_typeAttributesMap.Count == 0)
                    {
                        throw new System.Exception("远程属性（不同类型需要有哪些属性）未获得！");
                    }

                    string fileType = "DRW";
                    g_needAttributeNameList = GetNeedAttributeNameList(fileType, g_typeAttributesMap);
                }
            }
            catch (System.Exception ex)
            {
                logger.Error($"Exception in LoadProperties.{ex.Message}");
                throw new System.Exception($"Exception in LoadProperties.{ex.Message}");
            }
        }

        private void SetAppConfigValue(String paramName, String paramValue)
        {
            //logger.Info($"SetAppConfigValue,{paramName}:{paramValue}");
            ConfigManager.SetAppSetting(paramName, paramValue);
        }

        private void ReconnEnoviaIfDisconnected()
        {
            string responseJson;
            int iRet = connector.ExecuteMQL("version;", g_userName, g_password, "return", out responseJson);
            logger.Debug($"responseJson={responseJson}");
            if (!connector.IsEnoviaLoginByResponse(responseJson))
            {
                iRet = connector.EnoviaLogin(g_userName, g_password, "return", g_securityContext);
                if (iRet != 0)
                {
                    throw new Exception($"登录Enovia失败！{iRet}");
                }
            }
        }

        private void ExecuteTask()
        {
            //logger.Info("enter ExecuteTask...");
            string taskId = "";
            try
            {
                string responseJson = "", jParam = "";

                if (!IsCADRunning())
                {
                    try
                    {
                        session = StartCADApp(g_appPath);
                    }
                    catch (Exception ex)
                    {
                        string errMsg = "Failed to connect to Creo instance: " + ex.Message;
                        logger.Error(errMsg);
                        //MessageBox.Show(errMsg, "运行错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                ReconnEnoviaIfDisconnected();

                taskId = ConfigManager.GetAppSetting("TaskId"); //ConfigurationManager.AppSettings["TaskId"];
                if (taskId.Length > 0)
                {
                    logger.Info($"get Draw Id From Last Task. {taskId}");
                    jParam = "{\"program\":\"CASSConvertManage\",\"function\":\"getDrawIdByTaskId\",\"data\":[{\"objectId\":\""+ taskId + "\"}]}";
                    logger.Debug($"jParam={jParam}");
                    int iRet = connector.ExecuteJPO(jParam, g_userName, g_password, "return", out responseJson, g_securityContext);
                    if (iRet != 0)
                    {
                        throw new Exception("执行ExecuteJPO失败！" + responseJson);
                    }
                }
                else
                {
                    //logger.Info($"get Draw Id From Task List.");
                    jParam = "{\"program\":\"CASSConvertManage\",\"function\":\"getDrawIdFromTaskList\",\"data\":[{\"integName\":\"" + g_integName + "\"}]}";
                    logger.Debug($"jParam={jParam}");
                    int iRet = connector.ExecuteJPO(jParam, g_userName, g_password, "return", out responseJson, g_securityContext);
                    if (iRet != 0)
                    {
                        throw new Exception("执行ExecuteJPO失败1！" + responseJson);
                    }
                }

                logger.Debug($"responseJson1={responseJson}");
                string code = connector.getReturnCode(responseJson);
                if (code.Equals("0"))
                {
                    taskId = connector.getReturnSingleStringData(responseJson, "taskId");
                    string drawId = connector.getReturnSingleStringData(responseJson, "objectId");
                    string fileName = connector.getReturnSingleStringData(responseJson, "fileName");
                    string md5 = connector.getReturnSingleStringData(responseJson, "md5");

                    if (drawId.Trim().Length > 0 && taskId.Trim().Length > 0)
                    {
                        SetAppConfigValue("TaskId", taskId);
                        logger.Info($"drawId={drawId}, taskId={taskId}, fileName={fileName}.{openSWDocCount}");
                        timer1.Interval = iRunTime;//20250307
                        OpenWaitCount = 1;
                        //timerWaitOpen.Start();
                        backendTimer.Start();
                        ConvertMain(taskId, drawId, fileName, md5);
                        SetAppConfigValue("TaskId", "");
                    }
                    else
                    {
                        SetAppConfigValue("TaskId", "");
                        timer1.Interval = 30000;//20250307
                        logger.Info($"没有找到待处理任务！{openSWDocCount}");
                    }
                }
                else
                {
                    throw new Exception("执行失败！returnCode=" + code + ", returnMessage=" + connector.getReturnMessage(responseJson));
                }
            }
            catch (Exception ex)
            {
                logger.Error($"Exception in ExecuteTask. {ex.Message} -{taskId}");
                timer1.Interval = iRunTime;//20250309
            }
            finally
            {
                //timerWaitOpen.Stop();
                backendTimer.Stop();
            }
        }

        private void CreatePath(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        private string CreateTempFolder(string tmpFolder)
        {
            string tmpDestFullFolder = Path.Combine(Path.GetTempPath(), "EnoIntegCache");
            CreatePath(tmpDestFullFolder);
            tmpDestFullFolder = Path.Combine(tmpDestFullFolder, tmpFolder);
            CreatePath(tmpDestFullFolder);
            logger.Info($"CreateTempFolder: {tmpDestFullFolder}");
            return tmpDestFullFolder;
        }
        private void EmptyFolder(string dir)
        {
            foreach (string d in Directory.GetFileSystemEntries(dir))
            {
                if (File.Exists(d))
                {
                    try
                    {
                        FileInfo fi = new FileInfo(d);
                        if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                            fi.Attributes = FileAttributes.Normal;
                        File.Delete(d);//直接删除其中的文件 
                    }
                    catch
                    {

                    }
                }
                else
                {
                    try
                    {
                        DirectoryInfo d1 = new DirectoryInfo(d);
                        if (d1.GetFiles().Length != 0)
                        {
                            EmptyFolder(d1.FullName);////递归删除子文件夹
                        }
                        Directory.Delete(d);
                    }
                    catch
                    {

                    }
                }
            }
        }
        public void DeleteFolder(string dir)
        {
            foreach (string d in Directory.GetFileSystemEntries(dir))
            {
                if (File.Exists(d))
                {
                    FileInfo fi = new FileInfo(d);
                    if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                        fi.Attributes = FileAttributes.Normal;
                    File.Delete(d);//直接删除其中的文件 
                }
                else
                {
                    EmptyFolder(d);////递归删除子文件夹
                    Directory.Delete(d);
                }
            }
            Directory.Delete(dir);
        }

        public IpfcModel OpenCreoDoc(string fileFullName)
        {
            string filePath = Path.GetDirectoryName(fileFullName);
            string fileExt = Path.GetExtension(fileFullName).Substring(1);
            string fileName = Path.GetFileName(fileFullName);

            session.ChangeDirectory(filePath);
            CCpfcModelDescriptor descModelCreate = new CCpfcModelDescriptor();
            int fileType = 1;
            if (fileExt.ToUpper().Equals("PRT"))
                fileType = (int)EpfcModelType.EpfcMDL_PART;
            else if(fileExt.ToUpper().Equals("ASM"))
                fileType = (int)EpfcModelType.EpfcMDL_ASSEMBLY;
            else if (fileExt.ToUpper().Equals("DRW"))
                fileType = (int)EpfcModelType.EpfcMDL_DRAWING;
            IpfcModelDescriptor descModel = descModelCreate.Create(fileType, fileName, null);
            IpfcModel model = session.RetrieveModel(descModel); // 在Creo中显示模型
            model.Display();
            return model;
        }

        private void GenerateOutputFile(IpfcModel model, string outputFileName, string derivedOutputType)
        {
            try
            {
                logger.Info($"enter GenerateOutputFile, outputFileName={outputFileName}");

                if (derivedOutputType.Equals("jpg"))
                {
                    IpfcWindow window = session.get_CurrentWindow();
                    IpfcJPEGImageExportInstructions jpegInstrs;
                    CCpfcJPEGImageExportInstructions cjins = new CCpfcJPEGImageExportInstructions();
                    jpegInstrs = cjins.Create(10.0, 7.5);
                    IpfcRasterImageExportInstructions ins = (IpfcRasterImageExportInstructions)jpegInstrs;
                    ins.DotsPerInch = EpfcDotsPerInch.EpfcRASTERDPI_100;
                    ins.ImageDepth = EpfcRasterDepth.EpfcRASTERDEPTH_24;
                    window.ExportRasterImage(outputFileName, ins);
                }
                else
                {
                    //IpfcExportOptions exportOptions = (IpfcExportOptions)new CMpfcExportOptions();
                    //exportOptions.Type = (int)EpfcExportType.EpfcEXPORT_STEP;
                    int expType = 0;
                    if (derivedOutputType.Equals("stp"))
                    {
                        expType = (int)EpfcExportType.EpfcEXPORT_STEP;
                        model.ExportIntf3D(outputFileName, expType, null);
                    }
                    else if (derivedOutputType.Equals("igs"))
                    {
                        expType = (int)EpfcExportType.EpfcEXPORT_IGES_3D;
                        model.ExportIntf3D(outputFileName, expType, null);
                    }
                    else if (derivedOutputType.Equals("pdf"))
                    {
                        IpfcPDFExportInstructions pdfExportInstructions = new CCpfcPDFExportInstructions().Create();
                        pdfExportInstructions.ProfilePath = g_pdfProfilePath;
                        model.Export(outputFileName, (IpfcExportInstructions)pdfExportInstructions);
                    }
                    else if (derivedOutputType.Equals("dxf"))
                    {
                        IpfcDXFExportInstructions dxfExportInstructions = new CCpfcDXFExportInstructions().Create();
                        model.Export(outputFileName, (IpfcExportInstructions)dxfExportInstructions);
                    }
                }
            }
            catch (System.Exception ex)
            {
                string errmsg = $"Exception in GenerateOutputFile.{ex.Message}";
                throw new Exception(errmsg);
            }
        }
        private void ExportAFile(IpfcModel model, Dictionary<string, string> typeDerivedOutputsMap, string targetFileName, string fileType, JObject dataObj)
        {
            try
            {
                logger.Info($"enter ExportAFile...targetFileName={targetFileName}, fileType={fileType}");
                string needDerivedOutputs = typeDerivedOutputsMap[fileType];
                string[] needDerivedOutputList = needDerivedOutputs.Split(',');
                for (int k = 0; k < needDerivedOutputList.Length; k++)
                {
                    JObject jsDerivedOutput = new JObject();
                    string outputFileName = targetFileName + "." + needDerivedOutputList[k];
                    GenerateOutputFile(model, outputFileName, needDerivedOutputList[k]);
                    jsDerivedOutput["fileFullName"] = outputFileName;
                    ((JArray)dataObj["derivedOutputs"]).Add(jsDerivedOutput);
                }
            }
            catch (System.Exception ex)
            {
                string errmsg = $"Exception in ExportAFile.{ex.Message}";
                throw new Exception(errmsg);
            }
        }

        private void updateTaskState(string taskId, string taskState, string taskErr)
        {
            logger.Info("enter updateTaskState...");
            try
            {
                string jParam = "{\"program\":\"CASSConvertManage\",\"function\":\"updateTaskInfo\",\"data\":{\"taskId\": \"" + taskId +
                    "\",\"taskState\": \"" + taskState + "\",\"taskErr\": \"" + taskErr + "\"}}";
                logger.Debug($"jParam2={jParam}");
                string responseJson;
                int iRet = connector.ExecuteJPO(jParam, g_userName, g_password, "return", out responseJson);
                logger.Debug($"responseJson3={responseJson}");
                if (iRet == 0)
                {
                    logger.Info($"updateTaskState successful!");

                }
                else
                    throw new Exception($"updateTaskState failed! {iRet}");
            }
            catch (System.Exception ex)
            {
                string errmsg = $"Exception in updateTaskState.{ex.Message}";
                throw new Exception(errmsg);
            }
        }

        //private void GetDrawAttributeMap(ModelDoc2 swModel, string[] needAttributeNameList, JObject jsData)
        //{
        //    try
        //    {
        //        string val = "", valout = "";
        //        jsData["attributes"] = new JObject();

        //        bool wasResolved, linkToProp;
        //        int lRetVal;
        //        ModelDocExtension modelDocExtension = swModel.Extension;
        //        CustomPropertyManager swCustProps = modelDocExtension.get_CustomPropertyManager("");
        //        for (int i = 0; i < needAttributeNameList.Length; i++)
        //        {
        //            string attributeName = needAttributeNameList[i];
        //            logger.Info($"attributeName={attributeName}");
        //            string attributeValue = "";

        //            lRetVal = swCustProps.Get6(attributeName, false, out val, out valout, out wasResolved, out linkToProp);
        //            if (lRetVal != 1)
        //            {
        //                attributeValue = val;
        //            }
        //            logger.Info($"attributeName={attributeName}, attributeValue={attributeValue}");

        //            string attrNameText = MyUtils.GBK2UTF8(attributeName);
        //            string attrValueText = MyUtils.GBK2UTF8(attributeValue);
        //            jsData["attributes"][attrNameText] = attrValueText;
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        string errmsg = $"Exception in DrawAttributeMap. {ex.Message}";
        //        throw new Exception(errmsg);
        //    }
        //}
        private void ConvertMain(string taskId, string drawId, string fileName, string md5)
        {
            logger.Info("enter ConvertMain...");
            string tmpCacheFolder = "";
            string tmpOutputFolder = "";
            IpfcModel model = null;
            int maxRetries = 3;//add 20250303
            int retryCount = 0;
            bool isSuccess = false;
            while (retryCount < maxRetries && !isSuccess)//add 20250303
            {
                try
                {
                    if (session == null) throw new Exception($"session is null!");

                    //string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                    string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");//mod 20250303
                    string timestamp1 = null;
                    //tmpCacheFolder = CreateTempFolder(timestamp);
                    tmpCacheFolder = CreateTempFolder("ConvertCache");//20250301

                    JArray dataArray = new JArray();
                    JObject dataObj = new JObject();
                    dataObj.Add("objectId", drawId);
                    dataArray.Add(dataObj);

                    Stopwatch myStopwatch = new Stopwatch();
                    myStopwatch.Start();

                    //add 20250301
                    string fileFullName = Path.Combine(tmpCacheFolder, fileName);
                    int iRet = 0;
                    string jParam = null;
                    Boolean fileExist = false;
                    if (File.Exists(fileFullName))
                    {
                        string localMd5 = MyUtils.GetFileMd5Code(fileFullName);
                        if (localMd5.Equals(md5))
                        {
                            fileExist = true;
                        }
                    }

                    if (!fileExist)
                    {
                        logger.Info($"下载图纸！{fileName}");
                        jParam = "[{\"integName\": \"" + g_integName + "\",\"method\": \"download\",\"drwing\": \"yes\",\"localFolder\": \"" + tmpCacheFolder.Replace("\\", "\\\\") + "\",	\"data\": " + dataArray.ToString() + "}]";
                        logger.Debug($"jParam={jParam}");
                        iRet = connector.DownloadByTicket("CassPluginIntegrationJPO", "checkOutFilesCreateticket", jParam, g_userName, g_password, "return");
                        logger.Info("download finish!");
                        logger.Info($"Download {drawId} to {tmpCacheFolder}. Time = {myStopwatch.Elapsed.TotalSeconds}");
                    }

                    string remoteTmpFolder = "";
                    if (iRet == 0)
                    {
                        myStopwatch.Reset(); myStopwatch.Restart();
                        logger.Info("打开图纸！");
                        model = OpenCreoDoc(fileFullName);
                        if (model == null)
                        {
                            throw new Exception("openCreoDoc model is null!");
                        }
                        logger.Info($"OpenDoc Time = {myStopwatch.Elapsed.TotalSeconds}");

                        dataObj.Add("fileFullName", fileFullName);

                        timestamp1 = timestamp + "OP";
                        tmpOutputFolder = CreateTempFolder(timestamp1);
                        string fileNameWithExt = Path.GetFileName(fileFullName);
                        string fileNameNoExt = Path.GetFileNameWithoutExtension(fileFullName);
                        string targetFileName = Path.Combine(tmpOutputFolder, fileNameNoExt);
                        string fileType = Path.GetExtension(fileFullName).Substring(1).ToUpper();

                        myStopwatch.Reset(); myStopwatch.Restart();
                        logger.Info("生成衍生文件！");
                        dataObj["derivedOutputs"] = new JArray();
                        ExportAFile(model, g_typeDerivedOutputsMap, targetFileName, fileType, dataObj);
                        logger.Info($"Export Time = {myStopwatch.Elapsed.TotalSeconds}");

                        ReconnEnoviaIfDisconnected();

                        //20250211如果是工程图，进行属性映射
                        //if (fileType.Equals("DRW"))
                        //{
                        //    myStopwatch.Reset(); myStopwatch.Restart();
                        //    GetDrawAttributeMap(swModel, g_needAttributeNameList, dataObj);
                        //    logger.Info($"Attribute Map Time = {myStopwatch.Elapsed.TotalSeconds}");
                        //}

                        //myStopwatch.Reset(); myStopwatch.Restart();
                        //logger.Info("关闭文件！");
                        //swApp.CloseDoc(fileFullName);
                        //logger.Info($"Close Time = {myStopwatch.Elapsed.TotalSeconds}");

                        myStopwatch.Reset(); myStopwatch.Restart();
                        remoteTmpFolder = connector.getLoginUser() + "/" + timestamp1;
                        logger.Info($"tmpOutputFolder={tmpOutputFolder}, remoteTmpFolder={remoteTmpFolder}");
                        int ret;
                        ret = connector.UploadFolderToServer(tmpOutputFolder, remoteTmpFolder, true, false, true);
                        if (ret == 0)
                            logger.Info($"UploadFolderToServer completely! remoteTmpFolder={timestamp1}");
                        else
                            throw new Exception($"connector.UploadFolderToServer failed! {ret}");
                        logger.Info($"Upload Time = {myStopwatch.Elapsed.TotalSeconds}");
                    }
                    else
                    {
                        throw new Exception($"connector.DownloadByTicket failed! {iRet}");
                    }

                    ReconnEnoviaIfDisconnected();//add 20250307

                    myStopwatch.Reset(); myStopwatch.Restart();
                    jParam = "{\"program\":\"CASSConvertManage\",\"function\":\"AddDerivedOutput\",\"data\":{\"integName\": \"" + g_integName +
                        "\",\"tmpRemoteFolder\": \"" + timestamp1 + "\",	\"data\": " + dataArray.ToString() + "}}";
                    logger.Debug($"jParam1={jParam}");
                    string responseJson;
                    iRet = connector.ExecuteJPO(jParam, g_userName, g_password, "return", out responseJson);
                    logger.Debug($"responseJson2={responseJson}, iret={iRet}");
                    if (iRet == 0)
                    {
                        logger.Info($"Add Derived Output. Time = {myStopwatch.Elapsed.TotalSeconds}");
                        JObject resp = JObject.Parse(responseJson);
                        string code = (String)resp["code"];
                        if (code.Equals("0"))
                        {
                            logger.Info($"add Derived Output successful!");

                            //add 20250315
                            connector.DeleteDirectoryInServer(remoteTmpFolder);

                            myStopwatch.Reset(); myStopwatch.Restart();
                            updateTaskState(taskId, "0", "");
                            logger.Info($"UpdateState Time = {myStopwatch.Elapsed.TotalSeconds}");
                        }
                        else
                        {
                            string resultMesg = (String)resp["resultMesg"];
                            logger.Info($"resultMesg={resultMesg}");
                            throw new Exception($"AddDerivedOutput failed! {resultMesg}");
                        }

                        isSuccess = true;
                    }
                    else
                    {
                        throw new Exception($"AddDerivedOutput failed! {iRet}");
                    }


                }
                catch (Exception ex)
                {
                    retryCount++;
                    string errMsg = $"Exception in ConvertMain. {ex.Message.Replace("\r", "/").Replace("\n", "/")}";
                    logger.Error($"{errMsg} -{taskId} -{retryCount}");

                    //add 20250306
                    if (errMsg.Contains("Exception in GenerateOutputFile.") || errMsg.Contains("Exception in OpenSWDoc."))
                    {
                        RestartCADApp();
                    }

                    if (retryCount >= maxRetries) //add 20250303
                    {
                        updateTaskState(taskId, "1", errMsg);
                        // 达到最大重试次数，抛出异常或处理错误
                        throw new Exception("操作失败，已达最大重试次数！{errMsg}");
                    }

                    // 可选：添加重试间隔（如 1 秒）
                    Thread.Sleep(1000);
                }
                finally
                {
                    //if (!string.IsNullOrEmpty(tmpCacheFolder))
                    //{
                    //    if (Directory.Exists(tmpCacheFolder))
                    //    {
                    //        logger.Info($"DeleteFolder(tmpCacheFolder).{tmpCacheFolder}");
                    //        DeleteFolder(tmpCacheFolder);
                    //    }
                    //}
                    try
                    {
                        if (!string.IsNullOrEmpty(tmpOutputFolder))
                        {
                            if (Directory.Exists(tmpOutputFolder))
                            {
                                logger.Info($"DeleteFolder(tmpOutputFolder).{tmpOutputFolder}");
                                DeleteFolder(tmpOutputFolder);
                            }
                            string zipFileName = tmpOutputFolder + ".zip";
                            if (File.Exists(zipFileName))
                            {
                                File.Delete(zipFileName);
                            }
                        }
                        if (model != null)//add 20250301
                        {
                            session.get_CurrentWindow().Clear();
                        }
                    }
                    catch (Exception e)//20250303
                    {
                        logger.Error(e.Message);
                    }
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            ExecuteTask();
            timer1.Start();
        }
        private void MainForm_Activated(object sender, EventArgs e)
        {
            this.Visible = true;
        }

        //private void timerWaitOpen_Tick(object sender, EventArgs e)
        //{
        //    logger.Info($"enter timerWaitOpen_Tick...{OpenWaitCount}");
        //    if(OpenWaitCount % iOpenWaitMin == 0)
        //    {
        //        timerWaitOpen.Stop();
        //        KillConvertProcess();
        //    }
        //    OpenWaitCount++;
        //}
        private void BackendTimer_Elapsed(object sender, EventArgs e)
        {
            logger.Info($"enter BackendTimer_Elapsed...{OpenWaitCount}");
            if (OpenWaitCount % iOpenWaitMin == 0)
            {
                backendTimer.Stop();
                KillConvertProcess();
            }
            OpenWaitCount ++;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            backendTimer?.Dispose();
            base.OnFormClosing(e);
        }
    }
}
