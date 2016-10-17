using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Net;
using System.Security;
using ExcelService_tecWeb.Models;
using DocumentFormat.OpenXml;
using ClosedXML.Excel;
using TEC.VacacionesColaboradoresWeb;

namespace ExcelService_tecWeb
{
    public class DATA
    {
        public static ClientContext getContext()
        {
            ClientContext response = null;
            try
            {
                string user = "mitecsvcstest@svcs.itesm.mx";
                string password = "kj&saz9r";
                string siteUrl = ConfigurationManager.AppSettings["MiEspacionUrlTEST"].ToString();
                response = new ClientContext(siteUrl);
                //SecureString securePassword = new SecureString();
                //foreach (var c in password) { securePassword.AppendChar(c); }
                NetworkCredential credential = new NetworkCredential(user, password);
                //SharePointOnlineCredentials credential = new SharePointOnlineCredentials(user, securePassword);
                response.Credentials = credential;
            }
            catch(Exception ex)
            {
                ex.ToString();
            }
            return response;
        }
        public static void crearExcel(DATAExcel data)
        {
            Logger.getInstance().WriteLog("Generar Excel");
            try {
                var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("Registro2017");
                var Row = ws.Row(1);
                Logger.getInstance().WriteLog("creando titulos");
                for (int i = 0; i <  data.Titulos.Count; i++)
                {
                    Row.SetAutoFilter();
                    Row.Cell(i + 1).Value = data.Titulos[i];


                }
                Logger.getInstance().WriteLog("creando titulos END");
                Logger.getInstance().WriteLog("creando contenido");
                for (int i = 1; i < data.Contenido.Count; i++)
                {
                     var N_Row = ws.Row(i+1);
                     for (int c = 0; c < data.Contenido[i].Count; c++)
                     {
                         string it = data.Contenido[i-1][c].ToString();
                         N_Row.Cell(c+1).Value = it;
                     }
                         ////foreach (string it in data.Contenido[i - 1])
                         ////{
                            
                         ////}
                }
                Logger.getInstance().WriteLog("creando contenido END");

                    //var workbook = new XLWorkbook();
                    //var ws = workbook.Worksheets.Add("Registro2017");

                    //var column = ws.Column(1);
                    //column.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
                    //column.Cell(1).Value = "dato menso";
                    //columnFromWorksheet.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
                    //columnFromWorksheet.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
                    //columnFromWorksheet.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

                    //var columnFromRange = ws.Range("B1:B9").FirstColumn();

                    //columnFromRange.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
                    //columnFromRange.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
                    //columnFromRange.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
                    //columnFromRange.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

                    workbook.SaveAs(data.Path);
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }
        
    }
}