using System;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using Microsoft.Diagnostics.Tracing.Parsers;
using Microsoft.Diagnostics.Tracing.Parsers.Kernel;
using Microsoft.Diagnostics.Tracing.Session;
using System.IO.Ports;
using System.Drawing;
using System.Drawing.Imaging;
using System.Data;
using System.Reflection;
using System.Management;

namespace WindowsFormsApp1
{
    static class Program
    {       
        [STAThread]
        static void Main(string [] args)        
        {
            // Check that we're admin:
            if (!(TraceEventSession.IsElevated() ?? false))
            {
                MessageBox.Show("Please run as an Administrator....");
                return;
            }            
            // Start Watching process
            Thread watchProcess = new Thread(new ThreadStart(ProcessWatcher.starto));
            watchProcess.Start();
            ProcessWatcher.setProcessId(args[0]);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //Set up the form            
            Form1 f = new Form1();            
            Application.Run(f);
        }
    }
}
class ProcessWatcher
{
    static string ID;
    static string fileExtension = ".tif"; // change to .tif    
    public static void starto()
    {
        using (var kernelSession = new TraceEventSession("test"))
        {
            // Handle ctrl C :            
            Console.WriteLine("Setup cancel keys:");
            Console.CancelKeyPress += delegate (object sender, ConsoleCancelEventArgs e) {
                kernelSession.Dispose();
                Environment.Exit(0);
            }; 
            
            kernelSession.EnableKernelProvider(KernelTraceEventParser.Keywords.FileIO |
                                               KernelTraceEventParser.Keywords.FileIOInit |
                                               KernelTraceEventParser.Keywords.DiskFileIO);

            kernelSession.Source.Kernel.FileIOQueryInfo += fileCreate;
            // Start processing data:
            kernelSession.Source.Process();
        }
    }
    public static void killCameraApp()
    {
        KillProcessAndChildren(Int32.Parse(ID));
    }
    public static void setProcessId(String pID)
    {
        ProcessWatcher.ID = pID;
    }
    private static void fileCreate(FileIOInfoTraceData data)
    {
        if (data.ProcessID.ToString().Equals(ID) )
        {
            if (data.FileName.Contains(fileExtension))
            {
                FileIO.addCalDataToTiffFile(data.FileName);
            }
        }
    }
    private static void KillProcessAndChildren(int pid)
    {
        // Cannot close 'system idle process'.
        if (pid == 0)
        {            
            return;
        }
        Console.WriteLine("Trying to close: " + pid);
        System.Management.ManagementObjectSearcher searcher = new System.Management.ManagementObjectSearcher
                ("Select * From Win32_Process Where ParentProcessID=" + pid);
        System.Management.ManagementObjectCollection moc = searcher.Get();
        foreach (System.Management.ManagementObject mo in moc)
        {
            KillProcessAndChildren(Convert.ToInt32(mo["ProcessID"]));
        }
        try
        {            
            System.Diagnostics.Process proc = System.Diagnostics.Process.GetProcessById(pid);
            proc.Kill();
        }
        catch (ArgumentException)
        {
            // Process already exited.
        }
    }
}
class FileIO
{
    // using config text file for now --> will move to registry when building installer...
    private static String calPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\";
    private static String CONFIG = "turretWatcher.config";
    private static String calSetName;
    private static String currentCalFile;
    private static DataSet LUT;
    private static FileIO fileIO;
    private static bool WatcherMutex;
    private FileIO()
    {
        // Instantiate from config file...
        String configPath = calPath + CONFIG;
        String[] lines = System.IO.File.ReadAllLines(configPath);
        calSetName = lines[0];
        currentCalFile = calPath + calSetName + ".xml";
        initializeTurretObjectiveRelayLUT();
    }
    public static FileIO getInstance()
    {
        if(FileIO.fileIO == null)
        {
            FileIO.fileIO = new FileIO();
            return fileIO;
        }
        else
        {
            return fileIO;
        }
    }
    private static bool isFileLocked(string FileName)
    {
        FileStream fs = null;
        try
        {
            fs = File.Open(FileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (UnauthorizedAccessException) // https://msdn.microsoft.com/en-us/library/y973b725(v=vs.110).aspx
        {            
            try
            {
                fs = File.Open(FileName, FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (Exception)
            {
                return true; // This file has been locked, we can't even open it to read
            }
        }
        catch (Exception)
        {
            return true; // This file has been locked
        }
        finally
        {
            if (fs != null)
                fs.Close();
        }
        return false;
    }
    public static void addCalDataToTiffFile(string fName) // Will be called over and over... --> Need to handle many calls
    {
        if (WatcherMutex) return;        
        if (isFileLocked(fName)) return;
        WatcherMutex = true;
        try
        {                        
            Image img = Image.FromFile(fName);
            Image newImg = new Bitmap(img);
            System.Drawing.Imaging.PropertyItem[] items = img.PropertyItems;
            
            // Check that file hasn't already been stamped:
            bool hasValue = false;
            foreach (System.Drawing.Imaging.PropertyItem item in items)
            {
                if (item.Id == 6996) hasValue = true;
            }

            // ...if not, add stamp:
            if (!hasValue)
            {
                // Build string to write to file:
                string[] data = fileIO.parseFileNamePath(fName);
                if (data == null)
                {
                    try {
                        img.Dispose();
                        newImg.Dispose();
                        File.Delete(fName);
                       
                    }
                    catch
                    {

                    }
                    MessageBox.Show("Your file name " + fName + " is invalide.\n" +
                                  "Are you refering the corect lens/objective: (_5x1r, _1p25x1p6r, _100x2r,...)",
                                  "******* YOUR FILE HAS NOT BEEN SAVED!!!! **********",
                                  MessageBoxButtons.OK,
                                  MessageBoxIcon.Error,
                                  MessageBoxDefaultButton.Button3,
                                  MessageBoxOptions.DefaultDesktopOnly);
                    return;
                }

                string s = "\n[Calibration]\n" +
                           "Objective = " + data[0] + "\n" +
                           "Relay = " + data[1] + "\n" +
                           "PixelPitch = " + data[2] + "\n\n" +
                           LUT.Tables[0].Rows[0]["MicroscopeInfo"].ToString() + "\n";
                Console.WriteLine(s);
                char[] vs = s.ToCharArray();
                byte[] ba = new byte[vs.Length];
                for (int i = 0; i < ba.Length; i++)
                {
                    ba[i] = Convert.ToByte(vs[i]);
                }
                // Copy over tags from original image:
                for (int i = 0; i < items.Length; i++)
                {
                    newImg.SetPropertyItem(items[i]);
                }

                // Generate a 24bit RGB image with no compression:
                ImageCodecInfo myImageCodecInfo;
                Encoder myEncoder;
                Encoder myEncoderCol;
                EncoderParameter myEncoderParameter;
                EncoderParameters myEncoderParameters;
                myImageCodecInfo = GetEncoderInfo("image/tiff");
                myEncoder = Encoder.Compression;
                myEncoderCol = Encoder.ColorDepth;
                myEncoderParameters = new EncoderParameters(2);
                myEncoderParameter = new EncoderParameter(
                                    myEncoder,
                                    (long)EncoderValue.CompressionNone);
                myEncoderParameters.Param[0] = myEncoderParameter;
                myEncoderParameter = new EncoderParameter(myEncoderCol, 24L);
                myEncoderParameters.Param[1] = myEncoderParameter;

                // Build metadata for the image:
                System.Drawing.Imaging.PropertyItem item = img.PropertyItems[0];
                img.Dispose();
                item.Id = 6996;
                item.Len = vs.Length;
                item.Type = 2;
                item.Value = ba;
                newImg.SetPropertyItem(item);   
                
                //Save image:
                newImg.Save(fName, myImageCodecInfo, myEncoderParameters);
            }
            newImg.Dispose();
        } 
        catch (Exception e)
        {
            MessageBox.Show("Couldn't handle the file path passed by watcher. \n Exception: " + e.Message);
        }
        finally
        {
            WatcherMutex = false;
        }
    }
    private static ImageCodecInfo GetEncoderInfo(String mimeType)
    {
        int j;
        ImageCodecInfo[] encoders;
        encoders = ImageCodecInfo.GetImageEncoders();
        for (j = 0; j < encoders.Length; ++j)
        {
            if (encoders[j].MimeType == mimeType)
                return encoders[j];
        }
        return null;
    }
    public static void createNewTurretObjectiveRelayXML(DataSet ds)
    {        
        calSetName = "turretWatcher" + "_" + DateTimeOffset.Now.ToUnixTimeSeconds();
        currentCalFile = calPath + calSetName + ".xml";
        try
        {
            ds.WriteXml(currentCalFile);
        }
        catch(Exception e)
        {
            MessageBox.Show("Couldn't write to the new .xml file. \n Exception :" + e.Message);
            ProcessWatcher.killCameraApp();
            Environment.Exit(0);
        }
        updateConfigFile(calSetName);
        initializeTurretObjectiveRelayLUT();
    }
    public static void updateConfigFile(String newSetName)
    {
        String configPath = calPath + CONFIG;
        try
        {            
            StreamWriter sw = new StreamWriter(configPath);            
            sw.WriteLine(newSetName);
            sw.Close();
        }
        catch (Exception e)
        {
            MessageBox.Show("Couldn't read the turretWatcher.config file. \n Exception: " + e.Message);
            ProcessWatcher.killCameraApp();
            Environment.Exit(0);
        }        
    }

    public static string getCurrentCalibrationXML()
    {
        return currentCalFile;
    }    
    
    private static void initializeTurretObjectiveRelayLUT()
    {
        LUT = new DataSet(); // Tables[0] --> Microscope data, Tables[1] Objective/Relay - Pitch Combos (See .xml file)
        try
        {
            LUT.ReadXml(currentCalFile);
        }
        catch(FileNotFoundException e)
        {
            MessageBox.Show("Something went wrong reading the XML calibration file. You can: " +
                            "\n 1) Check that the turretWatcher.config file is refering to an \n   existing .xml file" +
                            "\n 2) Verify that the .xml file is correctly structured");
            ProcessWatcher.killCameraApp();
            Environment.Exit(0);
        }
        
    }
    public String[] getCalibration(String rO)
    {
        String[] data = { " ", " " , " " };
        DataTable tb = LUT.Tables[1];
        foreach (DataRow dr in tb.Rows)
        {
            if (dr["ID"].ToString().Equals(rO))
            {
                data[0] = dr["Objective"].ToString();
                data[1] = dr["Relay"].ToString();
                data[2] = dr["Pitch"].ToString();
            }
        }
        return data;
    }

    private String[] parseFileNamePath(String fName)
    {
        String sFName = System.IO.Path.GetFileNameWithoutExtension(fName)
                            .ToLower();
        String sfNameT = System.Text.RegularExpressions.Regex.Replace(sFName, @"\s", "");
        String sfNameTrim = System.Text.RegularExpressions.Regex.Replace(sfNameT, @"\.", "p");
        int index = sfNameTrim.LastIndexOf("_") + 1;

        if (index < 0) return null;
        String objRelay = sfNameTrim.Substring(index);

        String obj = null;
        String relay = null;        
        foreach (DataRow dr in LUT.Tables[1].Rows)
        {
            String compObj = System.Text.RegularExpressions.
                                    Regex.Replace(dr["Objective"].ToString() + "x", @"\.", "p");            
            if (objRelay.Contains(compObj)) {
                obj = dr["Objective"].ToString();
                break;
            }
        }

        if (obj == null) return null;

        String Relay = System.Text.RegularExpressions.Regex.Replace(objRelay, obj + "x", "");
        objRelay = Relay;
        
        if (!Relay.Contains("r") && Relay.Length > 0) objRelay = Relay + "r";
        if (Relay.Length == 0) objRelay = "1r";
        if (objRelay.Equals("1r")) relay = "1";        
        else if (objRelay.Equals("1p25r")) relay = "1.25";
        else if (objRelay.Equals("1p6r")) relay = "1.6";
        else if (objRelay.Equals("2r")) relay = "2";
        
        if (relay == null) return null;
        String[] data = { " ", " ", " " };
        DataTable tb = LUT.Tables[1];        
        foreach (DataRow dr in tb.Rows)
        {
            if (dr["Objective"].ToString().Equals(obj) && dr["Relay"].ToString().Equals(relay))
            {
                data[0] = dr["Objective"].ToString();
                data[1] = dr["Relay"].ToString();
                data[2] = dr["Pitch"].ToString();        
                return data;
            }
        }

        return null;
    }
}