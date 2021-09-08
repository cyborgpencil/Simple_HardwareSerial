// See https://aka.ms/new-console-template for more information
using HardwareSerial;
using HardwareSerial.Models;
using System.Management;
using System.Text;


// Create Objects
List<HardwareModel> hardwareModel = new List<HardwareModel>();
FileHandling fh = new FileHandling();


fh.ReadHostNamesFromTextFile();
//GetHardwareHostname();
//GetHardwareSerial();


//fh.AddListToWorksheet(hardwareModel);

void GetHardwareHostname()
{
    // HardwareModel Count based on the number of hostnames
    //hardwareModel = new List<HardwareModel>(fh.HostNames.Count);

    //Get the Hostnames
    foreach (var item in fh.HostNames)
    {
        HardwareModel  tempModel= new HardwareModel();
        tempModel.Hostname = item;
        hardwareModel.Add(tempModel);
    }
}

void GetHardwareSerial()
{
    for (int i = 0; i < hardwareModel.Count; i++)
    {
        try
        {
            ManagementScope scope = new ManagementScope($"\\\\{hardwareModel[i].Hostname}\\root\\cimv2");

            scope.Connect();

            ObjectQuery query = new ObjectQuery("SELECT SerialNumber, SerialNumber FROM Win32_BioS");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection serialInformation = searcher.Get();
            foreach (ManagementObject obj in serialInformation)
            {
                hardwareModel[i].BoardSerialName = obj["SerialNumber"].ToString();    
            }
            GetHardwareMonitorsModelName(hardwareModel[i], scope);
            searcher.Dispose();
        }
        catch (Exception ex)
        {
        }
    }
}

void GetHardwareMonitorsModelName(HardwareModel hw, ManagementScope scope)
{
    try
    {
        //Query ModelName
        string mModelQuery = "SELECT UserFriendlyName FROM WmiMonitorID ";
        ManagementObjectSearcher mModelSearcher = QueryObject($"\\\\{hw.Hostname}\\root\\WMI", mModelQuery);
        ManagementObjectCollection mModelCollection = mModelSearcher.Get();

        List<string> tempModelNames = new List<string>();


        foreach (ManagementObject m in mModelCollection)
        {
            hw.Monitors.Add(new MonitorInfo());

            foreach (PropertyData mData in m.Properties)
            {

                // Convert array to text
                UInt16[] x = (UInt16[])mData.Value;

                tempModelNames.Add(UInt16ArrayToString(x));
            }
        }

        // Send monitor info to grab Serial Number
        GetMonitorsSerialNumbers(hw, tempModelNames);


        mModelSearcher.Dispose();
    }
    catch (Exception)
    {  
    }
}

void GetMonitorsSerialNumbers(HardwareModel hw, List<string> tempModelNames)
{
    // Query Serial
    string mSerialQuery = "SELECT SerialNumberID FROM WmiMonitorID ";
    ManagementObjectSearcher mSerialSearcher = QueryObject($"\\\\{hw.Hostname}\\root\\WMI", mSerialQuery);
    ManagementObjectCollection mSerialCollection = mSerialSearcher.Get();

    List<string> tempSerialNumbers = new List<string>();
    foreach (ManagementObject s in mSerialCollection)
    {
        foreach (PropertyData sData in s.Properties)
        {
            // Convert array to text
            UInt16[] y = (UInt16[])sData.Value;

            tempSerialNumbers.Add(UInt16ArrayToString(y));
        }
    }

    for (int i = 0; i < hw.Monitors.Count; i++)
    {
        hw.Monitors[i].ModelName = tempModelNames[i];
        hw.Monitors[i].SerialNumber = tempSerialNumbers[i];
    }

    mSerialSearcher.Dispose();
}

// CONVERTER
string UInt16ArrayToString(UInt16[] input)
{
    // Final output
    StringBuilder output = new StringBuilder();

    // Turn UInt16 into Byte Array

    if (input != null)
    {
        foreach (ushort item in input)
        {
            byte[] b = BitConverter.GetBytes(item);

            //Encode each byte to string
            output.Append(Encoding.Default.GetString(b));
        }
    }
    return output.ToString();
}

// return Management object based on Query
ManagementObjectSearcher QueryObject(string scope, string query)
{

    try
    {
        // create a management scope object
        ManagementScope mScope = new ManagementScope(scope);
        mScope.Connect();

        //create object query
        ObjectQuery oQuery = new ObjectQuery(query);

        //create object searcher
        ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

        return searcher;
    }
    catch( Exception ex)
    {
        return null;
    }
}
