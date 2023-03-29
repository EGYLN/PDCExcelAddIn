using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Text;
using Excel=Microsoft.Office.Interop.Excel;

namespace PDCExcelCOMAddIn
{
    [Guid("5963B604-0BD8-4df9-AF32-B0A325430466")]

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]

    public class TestStruct
    {
        public int intValue;
        public string stringValue;
        public bool boolValue;
        public byte[] byteValues;
    }


    [Guid("A4FC4AF0-9D4C-4250-9760-BD2CD79567D3")]

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class PDC
    {
        public PDC()
        {
        }
        public static int StaticCallFunction()
        {
            return 666;
        }
        public TestStruct createTestStruct()
        {
            TestStruct tmpStruct = new TestStruct();
            tmpStruct.boolValue = true;
            tmpStruct.intValue = 39;
            tmpStruct.stringValue = "Hallo";
            tmpStruct.byteValues = new byte[] { 0, 1, 2, 3, 4 };
            return tmpStruct;
        }
        public TestStruct modifyTestStruct(TestStruct aStruct)
        {
            aStruct.boolValue = !aStruct.boolValue;
            aStruct.intValue++;
            aStruct.stringValue = "Reply";
            return aStruct;
        }
        public int InstanceCallFunction()
        {
            return 13;
        }
        public double MeinCall()
        {            
            return 0.34;
        }
       
        public void GetData(Excel.Range aTargetRange)
        {
            Excel.Worksheet tmpSheet = (Excel.Worksheet) aTargetRange.Parent;
            int tmpStartRow = aTargetRange.Row;
            int tmpStartColumn = aTargetRange.Column;
            object[,] tmpValues = new object[2,2];
            tmpValues[0,0] = "CompoundNo";
            tmpValues[1,0] = "Testno";
            tmpValues[0, 1] = "BAY 101079";
            tmpValues[1, 1] = "2588";
            Excel.Range tmpDataRange = tmpSheet.get_Range(
                tmpSheet.Cells[tmpStartRow, tmpStartColumn],
                tmpSheet.Cells[tmpStartRow+1, tmpStartColumn+1]);
            tmpDataRange.Value2 = tmpValues;
        }

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type, "Programmable"));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(GetSubKeyName(type, "InprocServer32"), true);
            key.SetValue("",System.Environment.SystemDirectory + @"\mscoree.dll",RegistryValueKind.String);
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type, "Programmable"), false);
        }

        private static string GetSubKeyName(Type type, string subKeyName)
        {
            System.Text.StringBuilder s = new System.Text.StringBuilder();
            s.Append(@"CLSID\{");
            s.Append(type.GUID.ToString().ToUpper());
            s.Append(@"}\");
            s.Append(subKeyName);
            return s.ToString();
        }  
    }
}
