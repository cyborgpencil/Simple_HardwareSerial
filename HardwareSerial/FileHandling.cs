﻿using HardwareSerial.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HardwareSerial
{
    public class FileHandling
    {
        public List<string> HostNames { get; set; }
        private string filename = @"Hardware Inventory.xlsx";
        private FileInfo ExcelFile { get; set; }
        ExcelWorksheet excelWorksheet { get; set; }
        ExcelPackage excel { get; set; }

        public FileHandling()
        {
            InitialExcelFileCreateAndSave();

            ExcelFile = new FileInfo(filename);
        }

        private void InitialExcelFileCreateAndSave()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            excel = new ExcelPackage();
            excel.Workbook.Worksheets.Add("Worksheet1");
            excelWorksheet = excel.Workbook.Worksheets["Worksheet1"];

            ExcelFile = new FileInfo(@"Hardware Inventory.xlsx");

            List<string[]> headerRow = new List<string[]>()
                {
                    new string[] {"HostName", "SerialNumber", "Monitor 1 ModelName", "Monitor 2 ModelName",
                    "Monitor 1 SerialNumber", "Monitor 2 SerialNumber"}
                };

            string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
            excelWorksheet.Cells[headerRange].LoadFromArrays(headerRow);

            excel.SaveAs(ExcelFile);

        }

        public void AddListToWorksheet(List<HardwareModel> hardwarelist)
        {
            for (int i = 0; i < hardwarelist.Count; i++)
            {

               excelWorksheet.Cells[i+2,1].LoadFromText(hardwarelist[i].Hostname);
               excelWorksheet.Cells[i + 2, 2].LoadFromText(hardwarelist[i].BoardSerialName);
                for (int j = 0; j < hardwarelist[i].Monitors.Count; j++)
                {
                    excelWorksheet.Cells[i + 2, 3+j].LoadFromText(hardwarelist[i].Monitors[j].ModelName);
                    excelWorksheet.Cells[i + 2, 5 + j].LoadFromText(hardwarelist[i].Monitors[j].SerialNumber);
                }
            }

            excel.SaveAs(ExcelFile);
        }

        public void ReadHostNamesFromTextFile()
        {
            HostNames = new List<string>();
            string[] lines = File.ReadAllLines(@"C:\hostnames\hostnames.txt");
            foreach (string line in lines)
            {
               HostNames.Add(line);
            }
        }

        public void HardwareModelToExcel(ExcelWorksheet eWorksheet, MonitorInfo mInfo)
        {

        }
    }
}