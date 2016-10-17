using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml;
using ClosedXML.Excel;
using ExcelService_tecWeb.Models;
using System.Configuration;
using System.Web.Configuration;
using ExcelService_tecWeb;
using TEC.VacacionesColaboradoresWeb;

namespace ExcelService_tecWeb.Controllers
{
    public class HomeController : Controller
    {
       // [SharePointContextFilter]
        
        public ActionResult Index()
        {
            Logger.getInstance().WriteLog(HttpContext.User.Identity.Name+"---UserPrincipal.Identity.IsAuthenticated:" + HttpContext.User.Identity.IsAuthenticated);
            try
            {

                if (ConfigurationManager.AppSettings["PATHSLN"] == null)
                {
                    var configuration = WebConfigurationManager.OpenWebConfiguration("~");
                    ExeConfigurationFileMap FileMap = new ExeConfigurationFileMap();
                    FileMap.ExeConfigFilename = Server.MapPath(@"~\Web.config");
                    Configuration Config = ConfigurationManager.OpenMappedExeConfiguration(FileMap, ConfigurationUserLevel.None);
                    Config.AppSettings.Settings.Add("PATHSLN", Server.MapPath(@"~\LOG"));
                    Config.Save(ConfigurationSaveMode.Modified);
                }
            }
            catch(Exception ex)
            {
                ex.ToString();
            }
            
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        [HttpGet]
        public virtual ActionResult DownloadExcel()
        {

            Logger.getInstance().WriteLog("DOWNLOAD");
            string fullPath = "";
            string file = "";
            string nombreLista = "";
            DATAExcel data = new DATAExcel();
            ClientContext ctx = null;
            try
            {
#if ADFSDebug
                Logger.getInstance().WriteLog("GETCONTEXT");
                System.Uri url = new System.Uri(ConfigurationManager.AppSettings["SITIO"].ToString());
                ctx = getContext(url);
                Logger.getInstance().WriteLog("GETCONTEXTEND");
               
#endif
#if DebugLocal
                Logger.getInstance().WriteLog("get context");
                ctx = DATA.getContext();
#endif         
                nombreLista = ConfigurationManager.AppSettings["NOMBRELISTA"].ToString();
                List Lista_Registro = ctx.Web.Lists.GetByTitle(nombreLista);
                CamlQuery q = new CamlQuery();
                ListItemCollection items = Lista_Registro.GetItems(q);
                FieldCollection fields = Lista_Registro.Fields;
                ctx.Load(items);
                ctx.Load(fields);               
                Logger.getInstance().WriteLog("ExecuteQuery");
                ctx.ExecuteQuery();
                Logger.getInstance().WriteLog("ExecuteQueryEnd:items");
                //FieldCollection fields = Lista_Registro.Fields;
                //ctx.Load(fields);
                //ctx.ExecuteQuery();
                int count = 0;
                List<string> Nombres_columna = new List<string>();
               // Nombres_columna.Add("Title");
                List<string> Internal_Names = new List<string>();
                Logger.getInstance().WriteLog("Calculo nombres:");
                foreach (Field field in fields)
                {

                    if (!field.FromBaseType||field.InternalName=="Title")
                    {
                        
                        string nombre = field.Title;
                        Nombres_columna.Add(nombre);
                        string internal_N = field.InternalName;
                        Internal_Names.Add(internal_N);
                    }
                }
                Logger.getInstance().WriteLog("Calculo nombres END:");
                List<List<string>> contenido = new List<List<string>>();
                Logger.getInstance().WriteLog("Calculo contenido:");
                foreach (ListItem item in items)
                {
                    List<string> stritem = new List<string>();
                    foreach (string f in Internal_Names)
                    {
                        //if (count == 40)
                        //    break;
                        string d = "";
                        try
                        {
                            d = item[f] != null ? item[f].ToString() : "";
                        }
                        catch (Exception ex)
                        {
                            ex.ToString();
                            d = "";
                        }
                        stritem.Add(d);
                        count += 1;
                    }
                    contenido.Add(stritem);
                    count = 0;

                }
                Logger.getInstance().WriteLog("Calculo contenido: END");
                data.Titulos = Nombres_columna;
                var itmd = items.ToList();
                ctx.Dispose();
                string filePath = Path.Combine(Server.MapPath("~/MyFiles/Registro.xlsx"));
                file = "Registro.xlsx";
                data.Path = filePath;
                data.Nombre = file;
                data.Contenido = contenido;
                DATA.crearExcel(data);
                fullPath = Path.Combine(Server.MapPath("~/MyFiles"), file);

            }
            catch (Exception ex)
            {
                Logger.getInstance().WriteLog(ex);
            }
            return File(fullPath, "application/vnd.ms-excel", file);
            //return File(workbook,"application/vnd.ms-excel","reportesotee.xlsx");
        }
        public ClientContext getContext(System.Uri url)
        {
            ClientContext response = null;
            try
            {

                var sharepointHostUrl = url;
                Logger.getInstance().WriteLog("URL:" + url + "--" + HttpContext.User.Identity.Name);
                response = TokenHelper.GetS2SClientContextWithClaimsIdentity(sharepointHostUrl, HttpContext.User, TokenHelper.IdentityClaimType.UPN, TokenHelper.ClaimProviderType.SAML, true);

            }
            catch (Exception ex)
            {
                Logger.getInstance().WriteLog(ex);
            }
            return response;
        }
    }
}
