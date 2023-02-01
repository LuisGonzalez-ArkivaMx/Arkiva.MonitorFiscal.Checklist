using System;
using System.IO;
using System.Runtime.Serialization;
using MFiles.VAF.Common;
using MFiles.VAF.Configuration;
using MFilesAPI;
using MFiles.VAF.Core;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Mail;
using System.Net.Mime;
using System.Linq;
using System.Threading;

namespace Arkiva.MonitorFiscal.Checklist
{    
    /// <summary>
    /// The entry point for this Vault Application Framework application.
    /// </summary>
    /// <remarks>Examples and further information available on the developer portal: http://developer.m-files.com/. </remarks>
    public partial class VaultApplication
        : ConfigurableVaultApplicationBase<Configuration>
    {
        CultureInfo defaultCulture;

        #region Overrides of VaultApplicationBase
        /// <inheritdoc />
        public override void StartOperations(Vault vaultPersistent)
        {
            string idioma = Configuration.ConfigurationServiciosGenerales.Idioma;
            defaultCulture = new CultureInfo(idioma); // set the desired culture here

            // you can also know the Client culture from env.CurrentUserSessionInfo.ClientCulture
            SysUtils.ReportInfoToEventLog("Configurando Idioma: " + defaultCulture.Name);
            Thread.CurrentThread.CurrentCulture = defaultCulture;
            Thread.CurrentThread.CurrentUICulture = defaultCulture;

            //Inicia Proceso de Validacion de Documentos
            this.StartBackgroundOperationTask();
            //this.StartBackgroundOperationTask_2();

            base.StartOperations(vaultPersistent);            
        }

        #endregion

        public static string LeerPlantilla(string ruta)
        {
            string line = "";
            using (StreamReader file = new StreamReader(ruta))
            {
                line = file.ReadToEnd();
                file.Close();
            }
            return line;
        }

        protected void StartBackgroundOperationTask()
        {
            SysUtils.ReportInfoToEventLog("Configurando Idioma: " + defaultCulture.Name);
            Thread.CurrentThread.CurrentCulture = defaultCulture;
            Thread.CurrentThread.CurrentUICulture = defaultCulture;

            try
            {
                this.BackgroundOperations.StartRecurringBackgroundOperation(
                    "Recurring Background Operation",
                    TimeSpan.FromMinutes(Configuration.ConfigurationServiciosGenerales.IntervaloDeEjecucionEnMins),
                () =>
                {
                    if (Configuration.ConfigurationServiciosGenerales.ApplicationEnabled.Equals("Yes"))
                    {
                        foreach (var grupo in Configuration.Grupos)
                        {
                            string sFechaActual = DateTime.Now.ToString("yyyy-MM-dd");

                            if (grupo.GroupEnabled.Equals("Yes") && ComparaFechaBaseContraUnaFechaInicioYFechaFin(sFechaActual, grupo.FechaInicio, grupo.FechaFin) == true)
                            {
                                // Inicializar objetos, clases y propiedades
                                var ot_ContactoExternoSE = PermanentVault
                                    .ObjectTypeOperations
                                    .GetObjectTypeIDByAlias("MF.OT.ExternalContact");

                                var pd_EstatusDocumento = PermanentVault
                                    .PropertyDefOperations
                                    .GetPropertyDefIDByAlias("PD.EstatusDocumento");

                                var pd_FrecuenciaDePagoDeNomina = PermanentVault
                                    .PropertyDefOperations
                                    .GetPropertyDefIDByAlias("PD.FrecuenciaDePagoDeNomina");

                                var pd_TipoDeValidacionLeyDeOutsourcing = PermanentVault
                                    .PropertyDefOperations
                                    .GetPropertyDefIDByAlias("PD.TipoDeValidacionLeyDeOutsourcing");

                                var pd_EstatusContactoExternoSE = PermanentVault
                                    .PropertyDefOperations
                                    .GetPropertyDefIDByAlias("PD.EstatusContactoExternoSE");

                                //var pd_ContactoExternoSE = PermanentVault
                                //    .PropertyDefOperations
                                //    .GetPropertyDefIDByAlias("PD.ContactoExternoServicioEspecializado.obj");

                                var pd_DocumentosSEContactoExterno = PermanentVault
                                    .PropertyDefOperations
                                    .GetPropertyDefIDByAlias("PD.DocumentosLeyOutsourcingContactoExterno");

                                var pd_EstatusProveedor = PermanentVault
                                    .PropertyDefOperations
                                    .GetPropertyDefIDByAlias("PD.EstatusProveedorLeyOutsourcing");

                                var pd_ValidacionManual = PermanentVault
                                    .PropertyDefOperations
                                    .GetPropertyDefIDByAlias("PD.ValidacionManual");

                                var pd_EstadoValidacionManual = PermanentVault
                                    .PropertyDefOperations
                                    .GetPropertyDefIDByAlias("PD.EstadoValidacionManual");

                                var searchBuilderOrganizacion = new MFSearchBuilder(PermanentVault);
                                searchBuilderOrganizacion.Deleted(false);

                                // Validar el filtro de busqueda de la Organizacion, si clase es null el filtro es por objeto
                                if (grupo.ValidacionOrganizacion.ClaseOrganizacion is null)
                                    searchBuilderOrganizacion.ObjType(grupo.ValidacionOrganizacion.ObjetoOrganizacion);
                                else // Si clase no es null, el filtro de busqueda es por clase
                                    searchBuilderOrganizacion.Class(grupo.ValidacionOrganizacion.ClaseOrganizacion);

                                foreach (var organizacion in searchBuilderOrganizacion.FindEx())
                                {
                                    SysUtils.ReportInfoToEventLog("Organizacion: " + organizacion.Title);

                                    bool bActivaProcesoChecklist = false;
                                    List<ObjVer> oListaDocumentosVigentes = new List<ObjVer>();
                                    bool bActivaRelacionDeDocumentosVigentes = false;
                                    List<ObjVer> oListaTodosLosDocumentosLO = new List<ObjVer>();
                                    List<ObjVerEx> contactosAdministradores = new List<ObjVerEx>();
                                    var oPropertyValues = new PropertyValues();
                                    bool bNotification = false;
                                    string sMainBodyMessage = "";
                                    string sBodyMessageDocuments = "";
                                    string sBodyMessageDocumentsEmp = "";
                                    string tBody = "";
                                    string tBodyEmpleado = "";

                                    string RutaPlantilla = @"C:\Notificacion\leyOutsourcing.html";// ConfigurationManager.AppSettings["RutaPlantilla"].ToString();
                                    string RutaTbody = @"C:\Notificacion\tbody.html";//ConfigurationManager.AppSettings["RutaTbody"].ToString();
                                    string RutaLista = @"C:\Notificacion\lista.html";//ConfigurationManager.AppSettings["RutaLista"].ToString();

                                    string RutaBanner = @"C:\Notificacion\img\Banner.png";// ConfigurationManager.AppSettings["RutaBanner"].ToString();
                                    string RutaCloud = @"C:\Notificacion\img\Cloud.png"; // ConfigurationManager.AppSettings["RutaCloud"].ToString();
                                    string RutaFooter = @"C:\Notificacion\img\Footer.png"; // ConfigurationManager.AppSettings["RutaFooter"].ToString();

                                    string RutaTbodyEmpleado = @"C:\Notificacion\tbodyEmpleado.html";
                                    string RutaListaEmpleado = @"C:\Notificacion\listaEmpleado.html";

                                    string Plantilla = LeerPlantilla(RutaPlantilla);

                                    oPropertyValues = PermanentVault
                                        .ObjectPropertyOperations
                                        .GetProperties(organizacion.ObjVer);

                                    if (oPropertyValues.IndexOf(grupo.ContactoAdministrador) != -1)
                                    {
                                        if (!oPropertyValues.SearchForPropertyEx(grupo.ContactoAdministrador, true).TypedValue.IsNULL())
                                        {
                                            contactosAdministradores = oPropertyValues
                                                .SearchForPropertyEx(grupo.ContactoAdministrador, true)
                                                .TypedValue
                                                .GetValueAsLookups().ToObjVerExs(PermanentVault);
                                        }
                                    }

                                    if (oPropertyValues.IndexOf(pd_TipoDeValidacionLeyDeOutsourcing) != -1 && //grupo.CheckboxLeyOutsourcing
                                        !oPropertyValues.SearchForPropertyEx(pd_TipoDeValidacionLeyDeOutsourcing, true).TypedValue.IsNULL()) //grupo.CheckboxLeyOutsourcing
                                    {
                                        bool bValidaDocumentoPorProyecto = false;
                                        List<ObjVerEx> searchResultsProyectoPorProveedor = new List<ObjVerEx>();

                                        var tipoValidacionLeyOutsourcing = oPropertyValues
                                            .SearchForPropertyEx(pd_TipoDeValidacionLeyDeOutsourcing, true)
                                            .TypedValue
                                            .GetLookupID();

                                        var nombreOTituloObjetoPadre = oPropertyValues
                                            .SearchForProperty((int)MFBuiltInPropertyDef
                                            .MFBuiltInPropertyDefNameOrTitle)
                                            .TypedValue
                                            .Value;

                                        var fechaInicioProveedor = oPropertyValues
                                            .SearchForPropertyEx(grupo.FechaInicioProveedor, true)
                                            .TypedValue
                                            .GetValueAsLocalizedText();

                                        // Validar tipo de proceso activado
                                        if (tipoValidacionLeyOutsourcing == 1) // Validacion por Proveedor
                                        {
                                            bActivaProcesoChecklist = true;

                                            SysUtils.ReportInfoToEventLog("Proceso 'Por Proveedor' activado: " + bActivaProcesoChecklist);
                                        }
                                        else if (tipoValidacionLeyOutsourcing == 2) // Validacion por Orden de Compra, Contrato y/o Proyecto
                                        {
                                            foreach (var documentoReferencia in grupo.ClasesReferencia)
                                            {
                                                var searchBuilderDocumentosReferencia = new MFSearchBuilder(PermanentVault);
                                                searchBuilderDocumentosReferencia.Deleted(false);
                                                searchBuilderDocumentosReferencia.Property
                                                (
                                                    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                                                    MFDataType.MFDatatypeLookup,
                                                    documentoReferencia.ClaseReferencia.ID
                                                );
                                                searchBuilderDocumentosReferencia.Property
                                                (
                                                    grupo.PropertyDefProveedorSEDocumentos,
                                                    MFDataType.MFDatatypeMultiSelectLookup,
                                                    organizacion.ObjVer.ID
                                                );

                                                var searchResultsDocumentosReferencia = searchBuilderDocumentosReferencia.FindEx();

                                                if (searchResultsDocumentosReferencia.Count > 0)
                                                {
                                                    foreach (var documento in searchResultsDocumentosReferencia)
                                                    {
                                                        var oProperties = documento.Properties;

                                                        var iEstatus = oProperties.SearchForPropertyEx(documentoReferencia.EstatusClaseReferencia.ID, true).TypedValue.GetLookupID();

                                                        if (documento.Class == 139) // Orden de Compra Emitida Proveedor
                                                        {
                                                            if (iEstatus == 1)
                                                            {
                                                                bActivaProcesoChecklist = true;
                                                                break;
                                                            }
                                                        }
                                                        else if (documento.Class == 113) // Contrato
                                                        {
                                                            if (iEstatus == 2 || iEstatus == 4)
                                                            {
                                                                bActivaProcesoChecklist = true;
                                                                break;
                                                            }
                                                        }
                                                        else if (documento.Class == 236) // Proyecto
                                                        {
                                                            if (iEstatus == 1)
                                                            {
                                                                var oLookupProyectoPorProveedor = new Lookup();
                                                                var oLookupsProyectoPorProveedor = new Lookups();

                                                                oLookupProyectoPorProveedor.Item = organizacion.ObjVer.ID;
                                                                oLookupsProyectoPorProveedor.Add(-1, oLookupProyectoPorProveedor);

                                                                // Buscar todos los proyectos del proveedor
                                                                var searchBuilderProyectoPorProveedor = new MFSearchBuilder(organizacion.Vault);
                                                                searchBuilderProyectoPorProveedor.Deleted(false);
                                                                searchBuilderProyectoPorProveedor.Class(documento.Class);
                                                                searchBuilderProyectoPorProveedor.Property
                                                                (
                                                                    grupo.PropertyDefProveedorSEDocumentos,
                                                                    MFDataType.MFDatatypeMultiSelectLookup,
                                                                    oLookupsProyectoPorProveedor
                                                                );

                                                                searchResultsProyectoPorProveedor = searchBuilderProyectoPorProveedor.FindEx();

                                                                bActivaProcesoChecklist = true;
                                                                bValidaDocumentoPorProyecto = true;

                                                                break;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            bActivaProcesoChecklist = false;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else if (tipoValidacionLeyOutsourcing == 3) // Validacion por Empresa Interna
                                        {
                                            bActivaProcesoChecklist = true;

                                            SysUtils.ReportInfoToEventLog("Proceso 'Por Empresa Interna' activado: " + bActivaProcesoChecklist);
                                        }

                                        if (bActivaProcesoChecklist == true)
                                        {
                                            SysUtils.ReportInfoToEventLog("Proveedor: " + organizacion.Title);

                                            string sChecklistDocumentName = "";
                                            bool bConcatenateDocument = false;
                                            string sPeriodoVencimientoDocumentoLO = "";
                                            string sDocumentosEnviadosAEmpleado = "";
                                            string sPeriodosDeDocumentosEnviadosAEmpleado = "";
                                            bool bDelete = true;

                                            // Eliminar del proveedor la informacion de documentos faltantes del recorrido anterior
                                            Sql.Query oQuery = new Sql.Query();
                                            oQuery.InsertarDocumentosFaltantesChecklist(bDelete, iProveedorID: organizacion.ObjVer.ID, sPeriodo: sFechaActual);

                                            // Recorrido de documentos proveedor
                                            foreach (var claseDocumento in grupo.DocumentosProveedor)
                                            {
                                                bDelete = false;

                                                List<ObjVer> oDocumentosVigentesPorValidar = new List<ObjVer>();
                                                List<ObjVer> oDocumentosVencidos = new List<ObjVer>();
                                                var sPeriodoDocumentoProveedor = "";
                                                var szNombreClaseDocumento = "";

                                                ObjectClass oObjectClass = PermanentVault
                                                    .ClassOperations
                                                    .GetObjectClass(claseDocumento.DocumentoProveedor.ID);

                                                szNombreClaseDocumento = oObjectClass.Name;

                                                DateTime dtFechaInicioPeriodo = Convert.ToDateTime(fechaInicioProveedor);

                                                DateTime dtFechaFinPeriodo = ObtenerRangoDePeriodoDelDocumento
                                                (
                                                    dtFechaInicioPeriodo,
                                                    claseDocumento.VigenciaDocumentoProveedor,
                                                    1
                                                );

                                                var searchBuilderDocumentosProveedor = new MFSearchBuilder(PermanentVault);
                                                searchBuilderDocumentosProveedor.Deleted(false);
                                                searchBuilderDocumentosProveedor.Property
                                                (
                                                    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                                                    MFDataType.MFDatatypeLookup,
                                                    claseDocumento.DocumentoProveedor.ID
                                                );
                                                searchBuilderDocumentosProveedor.Property
                                                (
                                                    grupo.PropertyDefProveedorSEDocumentos,
                                                    MFDataType.MFDatatypeMultiSelectLookup,
                                                    organizacion.ObjVer.ID
                                                );
                                                //searchBuilderDocumentosProveedor.Property
                                                //(
                                                //    MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow,
                                                //    MFDataType.MFDatatypeLookup,
                                                //    grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID
                                                //);
                                                //searchBuilderDocumentosProveedor.Property
                                                //(
                                                //    MFBuiltInPropertyDef.MFBuiltInPropertyDefState,
                                                //    MFDataType.MFDatatypeLookup,
                                                //    grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                //);

                                                if (searchBuilderDocumentosProveedor.FindEx().Count > 0) // Se encontro al menos un documento
                                                {
                                                    if (claseDocumento.VigenciaDocumentoProveedor != "No Aplica")
                                                    {
                                                        // Validar fecha fin contra la fecha actual                                            
                                                        string sFechaFin = dtFechaFinPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                                                        int iDateCompare = DateTime.Compare
                                                        (
                                                            Convert.ToDateTime(sFechaActual), // t1
                                                            Convert.ToDateTime(sFechaFin)     // t2
                                                        );

                                                        while (iDateCompare >= 0)
                                                        {
                                                            foreach (var documentoProveedor in searchBuilderDocumentosProveedor.FindEx())
                                                            {
                                                                SysUtils.ReportInfoToEventLog("Documento: " + documentoProveedor.Title + ", ID: " + documentoProveedor.ObjVer.ID);

                                                                bool bValidacionManual = false;
                                                                int classWorkflow = 0;

                                                                oPropertyValues = PermanentVault
                                                                    .ObjectPropertyOperations
                                                                    .GetProperties(documentoProveedor.ObjVer);

                                                                // Validar checkbox de validacion manual
                                                                if (oPropertyValues.IndexOf(pd_ValidacionManual) != -1)
                                                                {
                                                                    var validacionManual = Convert.ToBoolean(oPropertyValues.SearchForPropertyEx(pd_ValidacionManual, true).TypedValue.Value);

                                                                    if (validacionManual == true)
                                                                    {
                                                                        var estadoValidacionManual = oPropertyValues.SearchForPropertyEx(pd_EstadoValidacionManual, true).TypedValue.GetLookupID();

                                                                        if (estadoValidacionManual == 3) // Documento Validado
                                                                        {
                                                                            ActualizarWorkflowValidacionManual
                                                                            (
                                                                                documentoProveedor,
                                                                                grupo.ConfigurationWorkflow.WorkflowValidacionManual.WorkflowValidacionManualDocumento.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoValido.ID
                                                                            );
                                                                        }

                                                                        if (estadoValidacionManual == 4) // Documento No Validado
                                                                        {
                                                                            ActualizarWorkflowValidacionManual
                                                                            (
                                                                                documentoProveedor,
                                                                                grupo.ConfigurationWorkflow.WorkflowValidacionManual.WorkflowValidacionManualDocumento.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoNoValido.ID
                                                                            );
                                                                        }
                                                                    }
                                                                }

                                                                // Validar si el documento es para validacion manual
                                                                var iWorkflow = documentoProveedor
                                                                    .Vault
                                                                    .PropertyDefOperations
                                                                    .GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow)
                                                                    .ID;

                                                                classWorkflow = oPropertyValues.SearchForPropertyEx(iWorkflow, true).TypedValue.GetLookupID();

                                                                if (classWorkflow > 0)
                                                                {
                                                                    if (classWorkflow == grupo.ConfigurationWorkflow.WorkflowValidacionManual.WorkflowValidacionManualDocumento.ID)
                                                                    {
                                                                        bValidacionManual = true;
                                                                    }
                                                                }                                                                

                                                                // Obtener fecha del documento
                                                                var oFechaDeDocumento = oPropertyValues
                                                                    .SearchForPropertyEx(grupo.FechaDeDocumento, true)
                                                                    .TypedValue
                                                                    .Value;                                                                

                                                                DateTime dtFechaDeDocumento = Convert.ToDateTime(oFechaDeDocumento);

                                                                string sFechaDeDocumento = dtFechaDeDocumento.ToString("yyyy-MM-dd");

                                                                DateTime? dtFechaFinVigencia = null;

                                                                // Si existe la propiedad en la metadata del documento
                                                                if (oPropertyValues.IndexOf(grupo.FechaFinVigencia) != -1)
                                                                {
                                                                    if (!oPropertyValues.SearchForPropertyEx(grupo.FechaFinVigencia, true).TypedValue.IsNULL())
                                                                    {
                                                                        // Obtener fecha fin de vigencia
                                                                        var oFechaFinVigencia = oPropertyValues.SearchForPropertyEx(grupo.FechaFinVigencia, true).TypedValue.Value;
                                                                        dtFechaFinVigencia = Convert.ToDateTime(oFechaFinVigencia);
                                                                    }
                                                                }     
                                                                
                                                                if (bValidacionManual == false)
                                                                {
                                                                    if (claseDocumento.TipoValidacionVigenciaDocumento == "Por periodo")
                                                                    {
                                                                        // Validar si la fecha del documento esta dentro del periodo obtenido
                                                                        if (ComparaFechaBaseContraUnaFechaInicioYFechaFin(
                                                                            sFechaDeDocumento,
                                                                            dtFechaInicioPeriodo,
                                                                            dtFechaFinPeriodo) == true)
                                                                        {
                                                                            oDocumentosVigentesPorValidar.Add(documentoProveedor.ObjVer);

                                                                            // Actualizar el estatus "Vigente" al documento
                                                                            ActualizarEstatusDocumento
                                                                            (
                                                                                "Documento",
                                                                                documentoProveedor.ObjVer,
                                                                                pd_EstatusDocumento,
                                                                                1,
                                                                                grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.WorkflowValidacionesDocProveedor.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.EstadoDocumentoVigenteProveedor.ID,
                                                                                0,
                                                                                documentoProveedor,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                            );
                                                                        }
                                                                        else
                                                                        {
                                                                            oDocumentosVencidos.Add(documentoProveedor.ObjVer);

                                                                            // Agregar el estatus "Vencido" al documento
                                                                            ActualizarEstatusDocumento
                                                                            (
                                                                                "Documento",
                                                                                documentoProveedor.ObjVer,
                                                                                pd_EstatusDocumento,
                                                                                2,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoVencido.ID,
                                                                                0,
                                                                                documentoProveedor,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                            );
                                                                        }
                                                                    }
                                                                    else if (claseDocumento.TipoValidacionVigenciaDocumento == "Por fecha de vigencia")
                                                                    {
                                                                        // Validar si la fecha del documento esta dentro del periodo obtenido
                                                                        if (ComparaFechaBaseContraUnaFechaInicioYFechaFin(
                                                                            sFechaDeDocumento,
                                                                            dtFechaInicioPeriodo,
                                                                            dtFechaFinPeriodo) == true)
                                                                        {
                                                                            oDocumentosVigentesPorValidar.Add(documentoProveedor.ObjVer);
                                                                        }
                                                                        else
                                                                        {
                                                                            oDocumentosVencidos.Add(documentoProveedor.ObjVer);
                                                                        }

                                                                        // Validar vigencia tomando como referencia el ultimo periodo
                                                                        if (ValidarVigenciaDeDocumentoEnPeriodoActual(
                                                                            claseDocumento.VigenciaDocumentoProveedor,
                                                                            dtFechaDeDocumento,
                                                                            dtFechaFinVigencia) == true)
                                                                        {
                                                                            // Actualizar el estatus "Vigente" al documento
                                                                            ActualizarEstatusDocumento
                                                                            (
                                                                                "Documento",
                                                                                documentoProveedor.ObjVer,
                                                                                pd_EstatusDocumento,
                                                                                1,
                                                                                grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.WorkflowValidacionesDocProveedor.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.EstadoDocumentoVigenteProveedor.ID,
                                                                                0,
                                                                                documentoProveedor,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                            );
                                                                        }
                                                                        else
                                                                        {
                                                                            // Agregar el estatus "Vencido" al documento
                                                                            ActualizarEstatusDocumento
                                                                            (
                                                                                "Documento",
                                                                                documentoProveedor.ObjVer,
                                                                                pd_EstatusDocumento,
                                                                                2,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoVencido.ID,
                                                                                0,
                                                                                documentoProveedor,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                            );
                                                                        }
                                                                    }
                                                                }
                                                                else // Validacion manual es true
                                                                {
                                                                    var iState = documentoProveedor
                                                                        .Vault
                                                                        .PropertyDefOperations
                                                                        .GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefState)
                                                                        .ID;

                                                                    var classState = oPropertyValues.SearchForPropertyEx(iState, true).TypedValue.GetLookupID();

                                                                    if (classState == grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoNoValido.ID)
                                                                    {
                                                                        oDocumentosVencidos.Add(documentoProveedor.ObjVer);

                                                                        // Agregar el estatus "Vencido" al documento
                                                                        ActualizarEstatusDocumento
                                                                        (
                                                                            "Documento",
                                                                            documentoProveedor.ObjVer,
                                                                            pd_EstatusDocumento,
                                                                            2,
                                                                            grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                            grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoVencido.ID,
                                                                            0,
                                                                            documentoProveedor,
                                                                            grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                            grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                        );
                                                                    }

                                                                    if (classState == grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoValido.ID)
                                                                    {
                                                                        oDocumentosVigentesPorValidar.Add(documentoProveedor.ObjVer);

                                                                        // Actualizar el estatus "Vigente" al documento
                                                                        ActualizarEstatusDocumento
                                                                        (
                                                                            "Documento",
                                                                            documentoProveedor.ObjVer,
                                                                            pd_EstatusDocumento,
                                                                            1,
                                                                            grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.WorkflowValidacionesDocProveedor.ID,
                                                                            grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.EstadoDocumentoVigenteProveedor.ID,
                                                                            0,
                                                                            documentoProveedor,
                                                                            grupo.ConfigurationWorkflow.WorkflowValidacionManual.WorkflowValidacionManualDocumento.ID,
                                                                            grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoValido.ID
                                                                        );
                                                                    }
                                                                }                                                                

                                                                if (claseDocumento.TipoDocumentoChecklist == "Comprobante de pago")
                                                                {
                                                                    ObjVer oDocumentoRelacionado = new ObjVer();

                                                                    // Buscar documento relacionado al comprobante de pago
                                                                    if (oPropertyValues.IndexOf(claseDocumento.PropertyDefDocumentoRelacionado) != -1
                                                                        && !oPropertyValues.SearchForPropertyEx(claseDocumento.PropertyDefDocumentoRelacionado, true).TypedValue.IsNULL())
                                                                    {
                                                                        oDocumentoRelacionado = oPropertyValues
                                                                            .SearchForPropertyEx(claseDocumento.PropertyDefDocumentoRelacionado, true)
                                                                            .TypedValue
                                                                            .GetValueAsLookup().GetAsObjVer();
                                                                    }

                                                                    // Enviar informacion a la tabla de comprobantes de pago
                                                                    oQuery.InsertarComprobantesPago(
                                                                        szNombreClaseDocumento,
                                                                        documentoProveedor.ID,
                                                                        nombreOTituloObjetoPadre.ToString(),
                                                                        organizacion.ID,
                                                                        oDocumentoRelacionado.ID);
                                                                }
                                                            }

                                                            var sPeriodoDocumentoFaltante = ObtenerPeriodoDeDocumentoFaltante
                                                                (
                                                                    claseDocumento.VigenciaDocumentoProveedor,
                                                                    dtFechaInicioPeriodo,
                                                                    dtFechaFinPeriodo
                                                                );

                                                            if (oDocumentosVigentesPorValidar.Count < 1)
                                                            {
                                                                // Se agrega al correo el nombre del documento no encontrado
                                                                // en el periodo validado
                                                                string ListaItems = LeerPlantilla(RutaLista);
                                                                sChecklistDocumentName += ListaItems.Replace("[Documento]", szNombreClaseDocumento);

                                                                sPeriodoDocumentoProveedor = sPeriodoDocumentoFaltante;

                                                                // Modificar formato de la fecha del periodo del documento
                                                                if (claseDocumento.VigenciaDocumentoProveedor == "Mensual" ||
                                                                    claseDocumento.VigenciaDocumentoProveedor == "Bimestral" ||
                                                                    claseDocumento.VigenciaDocumentoProveedor == "Trimestral" ||
                                                                    claseDocumento.VigenciaDocumentoProveedor == "Cuatrimestral" ||
                                                                    claseDocumento.VigenciaDocumentoProveedor == "Anual")
                                                                {
                                                                    sPeriodoDocumentoProveedor = dtFechaInicioPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                }

                                                                if (claseDocumento.TipoDocumentoChecklist == "Documento checklist")
                                                                {

                                                                    // Insert
                                                                    oQuery.InsertarDocumentosFaltantesChecklist(
                                                                        bDelete,
                                                                        sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                        iProveedorID: organizacion.ObjVer.ID,
                                                                        sCategoria: "Documento Vencido",
                                                                        sTipoDocumento: "Documento Proveedor",
                                                                        sNombreDocumento: szNombreClaseDocumento,
                                                                        iDocumentoID: claseDocumento.DocumentoProveedor.ID,
                                                                        iClaseID: claseDocumento.DocumentoProveedor.ID,
                                                                        sVigencia: claseDocumento.VigenciaDocumentoProveedor,
                                                                        sPeriodo: sPeriodoDocumentoProveedor);
                                                                }

                                                                sPeriodoVencimientoDocumentoLO += sPeriodoDocumentoProveedor + "<br/>";

                                                                bNotification = true;
                                                                bConcatenateDocument = true;
                                                            }

                                                            // Enviar los documentos vencidos a la tabla de documentos faltantes checklist 
                                                            foreach (var documento in oDocumentosVencidos)
                                                            {
                                                                sPeriodoDocumentoProveedor = sPeriodoDocumentoFaltante;

                                                                // Modificar formato de la fecha del periodo del documento
                                                                if (claseDocumento.VigenciaDocumentoProveedor == "Mensual" ||
                                                                    claseDocumento.VigenciaDocumentoProveedor == "Bimestral" ||
                                                                    claseDocumento.VigenciaDocumentoProveedor == "Trimestral" ||
                                                                    claseDocumento.VigenciaDocumentoProveedor == "Cuatrimestral" ||
                                                                    claseDocumento.VigenciaDocumentoProveedor == "Anual")
                                                                {
                                                                    sPeriodoDocumentoProveedor = dtFechaInicioPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                }

                                                                if (claseDocumento.TipoDocumentoChecklist == "Documento checklist")
                                                                {
                                                                    SysUtils.ReportInfoToEventLog("Insertando documento: " + documento.ID + " en la tabla DocumentosCaducados");

                                                                    if (bValidaDocumentoPorProyecto == true)
                                                                    {
                                                                        foreach (var proyecto in searchResultsProyectoPorProveedor)
                                                                        {
                                                                            var pd_DocumentosRelacionadosAlProyecto = proyecto.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Document");
                                                                            var oPropertiesProyecto = proyecto.Properties;

                                                                            var oListDocumentosProyecto = oPropertiesProyecto
                                                                                .SearchForPropertyEx(pd_DocumentosRelacionadosAlProyecto, true)
                                                                                .TypedValue
                                                                                .GetValueAsLookups()
                                                                                .ToObjVerExs(proyecto.Vault);

                                                                            foreach (var documentoProyecto in oListDocumentosProyecto)
                                                                            {
                                                                                if (documento.ID == documentoProyecto.ObjVer.ID)
                                                                                {
                                                                                    // Insert la tabla de documentos caducados
                                                                                    oQuery.InsertarDocumentosCaducados(
                                                                                        sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                                        iProveedorID: organizacion.ObjVer.ID,
                                                                                        sProyecto: proyecto.Title,
                                                                                        iProyectoID: proyecto.ID,
                                                                                        sCategoria: "Documento Vencido",
                                                                                        sTipoDocumento: "Documento Proveedor",
                                                                                        sNombreDocumento: szNombreClaseDocumento,
                                                                                        iDocumentoID: documento.ID,
                                                                                        iClaseID: claseDocumento.DocumentoProveedor.ID,
                                                                                        sVigencia: claseDocumento.VigenciaDocumentoProveedor,
                                                                                        sPeriodo: sPeriodoDocumentoProveedor);
                                                                                }
                                                                                else
                                                                                {
                                                                                    // Enviar a la tabla de documentos faltantes
                                                                                    oQuery.InsertarDocumentosFaltantesChecklist(
                                                                                        bDelete,
                                                                                        sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                                        iProveedorID: organizacion.ObjVer.ID,
                                                                                        sProyecto: proyecto.Title,
                                                                                        iProyectoID: proyecto.ID,
                                                                                        sCategoria: "Documento Faltante",
                                                                                        sTipoDocumento: "Documento Proveedor",
                                                                                        sNombreDocumento: szNombreClaseDocumento,
                                                                                        iDocumentoID: 0,
                                                                                        iClaseID: claseDocumento.DocumentoProveedor.ID,
                                                                                        sVigencia: claseDocumento.VigenciaDocumentoProveedor,
                                                                                        sPeriodo: sPeriodoDocumentoProveedor);
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        // Insert
                                                                        oQuery.InsertarDocumentosCaducados(
                                                                            sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                            iProveedorID: organizacion.ObjVer.ID,
                                                                            sCategoria: "Documento Vencido",
                                                                            sTipoDocumento: "Documento Proveedor",
                                                                            sNombreDocumento: szNombreClaseDocumento,
                                                                            iDocumentoID: documento.ID,
                                                                            iClaseID: claseDocumento.DocumentoProveedor.ID,
                                                                            sVigencia: claseDocumento.VigenciaDocumentoProveedor,
                                                                            sPeriodo: sPeriodoDocumentoProveedor);
                                                                    }
                                                                }
                                                            }

                                                            // Crear nuevo periodo a partir de fecha fin que se convierte en la nueva fecha inicio
                                                            dtFechaInicioPeriodo = dtFechaFinPeriodo;
                                                            dtFechaFinPeriodo = ObtenerRangoDePeriodoDelDocumento
                                                            (
                                                                dtFechaInicioPeriodo,
                                                                claseDocumento.VigenciaDocumentoProveedor,
                                                                1
                                                            );

                                                            sFechaFin = "";
                                                            sFechaFin = dtFechaFinPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                                                            iDateCompare = DateTime.Compare
                                                            (
                                                                Convert.ToDateTime(sFechaActual), // t1
                                                                Convert.ToDateTime(sFechaFin)     // t2
                                                            );
                                                        }

                                                        if (oDocumentosVigentesPorValidar.Count > 0)
                                                        {
                                                            var sComparaFecha1 = "";
                                                            var sComparaFecha2 = "";
                                                            var dtFecha1 = new DateTime();
                                                            var dtFecha2 = new DateTime();
                                                            var dtFechaFinal = new DateTime();
                                                            var objVerDocumento1 = new ObjVer();
                                                            var objVerDocumento2 = new ObjVer();
                                                            var objVerDocumentoFinal = new ObjVer();

                                                            // Obtener el documento mas recientes de los encontrados
                                                            foreach (ObjVer documento in oDocumentosVigentesPorValidar)
                                                            {
                                                                oListaTodosLosDocumentosLO.Add(documento);

                                                                oPropertyValues = PermanentVault
                                                                    .ObjectPropertyOperations
                                                                    .GetProperties(documento);

                                                                if (oPropertyValues.IndexOf(pd_EstatusDocumento) != -1)
                                                                {
                                                                    var oFechaDocumento = oPropertyValues
                                                                    .SearchForPropertyEx(grupo.FechaDeDocumento, true)
                                                                    .TypedValue
                                                                    .Value;

                                                                    // Comparar fecha de documentos (misma clase de documento) encontrados
                                                                    // Obtener el documento mas reciente y vigente
                                                                    if (sComparaFecha1 == "")
                                                                    {
                                                                        dtFecha1 = Convert.ToDateTime(oFechaDocumento);
                                                                        sComparaFecha1 = dtFecha1.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                        objVerDocumento1 = documento;
                                                                    }
                                                                    else
                                                                    {
                                                                        dtFecha2 = Convert.ToDateTime(oFechaDocumento);
                                                                        sComparaFecha2 = dtFecha2.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                        objVerDocumento2 = documento;
                                                                    }

                                                                    if (sComparaFecha1 != "" && sComparaFecha2 != "")
                                                                    {
                                                                        int iComparaFechasDeDocumentoChecklist = DateTime.Compare
                                                                        (
                                                                            Convert.ToDateTime(sComparaFecha1),
                                                                            Convert.ToDateTime(sComparaFecha2)
                                                                        );

                                                                        if (iComparaFechasDeDocumentoChecklist < 0)
                                                                        {
                                                                            sComparaFecha1 = "";
                                                                            objVerDocumentoFinal = objVerDocumento2;
                                                                            dtFechaFinal = dtFecha2;
                                                                        }
                                                                        else
                                                                        {
                                                                            sComparaFecha2 = "";
                                                                            objVerDocumentoFinal = objVerDocumento1;
                                                                            dtFechaFinal = dtFecha1;
                                                                        }
                                                                    }

                                                                    // Si solo hay un documento, se establece el ID de objeto, la fecha de documento y la vigencia
                                                                    // directamente en el metodo ""
                                                                    if (oDocumentosVigentesPorValidar.Count == 1)
                                                                    {
                                                                        objVerDocumentoFinal = objVerDocumento1;
                                                                        dtFechaFinal = dtFecha1;
                                                                    }
                                                                }
                                                            }

                                                            // Se agrega a la lista el documento vigente
                                                            oListaDocumentosVigentes.Add(objVerDocumentoFinal);

                                                            // Activa la relacion de objetos
                                                            bActivaRelacionDeDocumentosVigentes = true;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        // Validar los documentos cuya vigencia es "No Aplica"                                                   
                                                        List<ObjVerEx> sbDocumentosProveedorNoAplicaVigencia = new List<ObjVerEx>();
                                                        sbDocumentosProveedorNoAplicaVigencia = searchBuilderDocumentosProveedor.FindEx();

                                                        // Agregar a lista de documentos vigentes
                                                        oListaDocumentosVigentes.Add(sbDocumentosProveedorNoAplicaVigencia[0].ObjVer);

                                                        // Activa la relacion de objetos
                                                        bActivaRelacionDeDocumentosVigentes = true;

                                                        foreach (ObjVerEx documentoProveedor in sbDocumentosProveedorNoAplicaVigencia)
                                                        {
                                                            bool bValidacionManual = false;
                                                            int classWorkflow = 0;

                                                            oPropertyValues = PermanentVault
                                                                .ObjectPropertyOperations
                                                                .GetProperties(documentoProveedor.ObjVer);

                                                            // Validar checkbox de validacion manual
                                                            if (oPropertyValues.IndexOf(pd_ValidacionManual) != -1)
                                                            {
                                                                var validacionManual = Convert.ToBoolean(oPropertyValues.SearchForPropertyEx(pd_ValidacionManual, true).TypedValue.Value);

                                                                if (validacionManual == true)
                                                                {
                                                                    var estadoValidacionManual = oPropertyValues.SearchForPropertyEx(pd_EstadoValidacionManual, true).TypedValue.GetLookupID();

                                                                    if (estadoValidacionManual == 3) // Documento Validado
                                                                    {
                                                                        ActualizarWorkflowValidacionManual
                                                                        (
                                                                            documentoProveedor,
                                                                            grupo.ConfigurationWorkflow.WorkflowValidacionManual.WorkflowValidacionManualDocumento.ID,
                                                                            grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoValido.ID
                                                                        );
                                                                    }

                                                                    if (estadoValidacionManual == 4) // Documento No Validado
                                                                    {
                                                                        ActualizarWorkflowValidacionManual
                                                                        (
                                                                            documentoProveedor,
                                                                            grupo.ConfigurationWorkflow.WorkflowValidacionManual.WorkflowValidacionManualDocumento.ID,
                                                                            grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoNoValido.ID
                                                                        );
                                                                    }
                                                                }
                                                            }

                                                            // Validar si el documento es para validacion manual
                                                            var iWorkflow = documentoProveedor
                                                                .Vault
                                                                .PropertyDefOperations
                                                                .GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow)
                                                                .ID;

                                                            classWorkflow = oPropertyValues.SearchForPropertyEx(iWorkflow, true).TypedValue.GetLookupID();

                                                            if (classWorkflow > 0)
                                                            {
                                                                if (classWorkflow == grupo.ConfigurationWorkflow.WorkflowValidacionManual.WorkflowValidacionManualDocumento.ID)
                                                                {
                                                                    bValidacionManual = true;
                                                                }
                                                            }

                                                            if (bValidacionManual == false)
                                                            {
                                                                // Agregar el estatus "Vigente" al documento
                                                                ActualizarEstatusDocumento
                                                                (
                                                                    "Documento",
                                                                    documentoProveedor.ObjVer,
                                                                    pd_EstatusDocumento,
                                                                    1,
                                                                    grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.WorkflowValidacionesDocProveedor.ID,
                                                                    grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.EstadoDocumentoVigenteProveedor.ID,
                                                                    0,
                                                                    documentoProveedor,
                                                                    grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                    grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                );
                                                            }
                                                            else // Validacion manual es true
                                                            {
                                                                var iState = documentoProveedor
                                                                    .Vault
                                                                    .PropertyDefOperations
                                                                    .GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefState)
                                                                    .ID;

                                                                var classState = oPropertyValues.SearchForPropertyEx(iState, true).TypedValue.GetLookupID();

                                                                if (classState == grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoNoValido.ID)
                                                                {
                                                                    oDocumentosVencidos.Add(documentoProveedor.ObjVer);

                                                                    // Agregar el estatus "Vencido" al documento
                                                                    ActualizarEstatusDocumento
                                                                    (
                                                                        "Documento",
                                                                        documentoProveedor.ObjVer,
                                                                        pd_EstatusDocumento,
                                                                        2,
                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoVencido.ID,
                                                                        0,
                                                                        documentoProveedor,
                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                    );
                                                                }

                                                                if (classState == grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoValido.ID)
                                                                {
                                                                    oDocumentosVigentesPorValidar.Add(documentoProveedor.ObjVer);

                                                                    // Actualizar el estatus "Vigente" al documento
                                                                    ActualizarEstatusDocumento
                                                                    (
                                                                        "Documento",
                                                                        documentoProveedor.ObjVer,
                                                                        pd_EstatusDocumento,
                                                                        1,
                                                                        grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.WorkflowValidacionesDocProveedor.ID,
                                                                        grupo.ConfigurationWorkflow.WorkflowDocumentoProveedor.EstadoDocumentoVigenteProveedor.ID,
                                                                        0,
                                                                        documentoProveedor,
                                                                        grupo.ConfigurationWorkflow.WorkflowValidacionManual.WorkflowValidacionManualDocumento.ID,
                                                                        grupo.ConfigurationWorkflow.WorkflowValidacionManual.EstadoDocumentoValido.ID
                                                                    );
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                else // No se encontro ningun documento de la clase validada
                                                {
                                                    if (claseDocumento.VigenciaDocumentoProveedor == "No Aplica")
                                                    {
                                                        // Si no se encuentra ningun documento de la clase de documento buscada
                                                        // Se agrega la clase de documento en el correo para que se suba a la boveda
                                                        string ListaItems = LeerPlantilla(RutaLista);
                                                        sChecklistDocumentName += ListaItems.Replace("[Documento]", szNombreClaseDocumento);

                                                        sPeriodoDocumentoProveedor = "Faltante";

                                                        sPeriodoVencimientoDocumentoLO += sPeriodoDocumentoProveedor + "<br/>";

                                                        if (claseDocumento.TipoDocumentoChecklist == "Documento checklist")
                                                        {
                                                            if (bValidaDocumentoPorProyecto == true)
                                                            {
                                                                foreach (var proyecto in searchResultsProyectoPorProveedor)
                                                                {
                                                                    // Enviar la informacion del documento faltante a la BD
                                                                    oQuery.InsertarDocumentosFaltantesChecklist(
                                                                        bDelete,
                                                                        sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                        iProveedorID: organizacion.ObjVer.ID,
                                                                        sProyecto: proyecto.Title,
                                                                        iProyectoID: proyecto.ID,
                                                                        sCategoria: "Documento Faltante",
                                                                        sTipoDocumento: "Documento Proveedor",
                                                                        sNombreDocumento: szNombreClaseDocumento,
                                                                        iDocumentoID: 0,
                                                                        iClaseID: claseDocumento.DocumentoProveedor.ID,
                                                                        sVigencia: claseDocumento.VigenciaDocumentoProveedor,
                                                                        sPeriodo: sPeriodoDocumentoProveedor);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                // Enviar la informacion del documento faltante a la BD
                                                                oQuery.InsertarDocumentosFaltantesChecklist(
                                                                    bDelete,
                                                                    sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                    iProveedorID: organizacion.ObjVer.ID,
                                                                    sCategoria: "Documento Faltante",
                                                                    sTipoDocumento: "Documento Proveedor",
                                                                    sNombreDocumento: szNombreClaseDocumento,
                                                                    iDocumentoID: 0,
                                                                    iClaseID: claseDocumento.DocumentoProveedor.ID,
                                                                    sVigencia: claseDocumento.VigenciaDocumentoProveedor,
                                                                    sPeriodo: sPeriodoDocumentoProveedor);
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        // Validar fecha fin contra la fecha actual                                            
                                                        string sFechaFin = dtFechaFinPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                                                        int iDateCompare = DateTime.Compare
                                                        (
                                                            Convert.ToDateTime(sFechaActual), // t1
                                                            Convert.ToDateTime(sFechaFin)     // t2
                                                        );

                                                        while (iDateCompare >= 0)
                                                        {
                                                            // Si no se encuentra ningun documento de la clase de documento buscada
                                                            // Se agrega la clase de documento en el correo para que se suba a la boveda
                                                            string ListaItems = LeerPlantilla(RutaLista);
                                                            sChecklistDocumentName += ListaItems.Replace("[Documento]", szNombreClaseDocumento);

                                                            var sPeriodoDocumentoFaltante = ObtenerPeriodoDeDocumentoFaltante
                                                            (
                                                                claseDocumento.VigenciaDocumentoProveedor,
                                                                dtFechaInicioPeriodo,
                                                                dtFechaFinPeriodo
                                                            );

                                                            sPeriodoDocumentoProveedor = sPeriodoDocumentoFaltante;

                                                            sPeriodoVencimientoDocumentoLO += sPeriodoDocumentoProveedor + "<br/>";

                                                            // Modificar formato de la fecha del periodo para enviarla a la BD
                                                            if (claseDocumento.VigenciaDocumentoProveedor == "Mensual" ||
                                                                claseDocumento.VigenciaDocumentoProveedor == "Bimestral" ||
                                                                claseDocumento.VigenciaDocumentoProveedor == "Trimestral" ||
                                                                claseDocumento.VigenciaDocumentoProveedor == "Cuatrimestral" ||
                                                                claseDocumento.VigenciaDocumentoProveedor == "Anual")
                                                            {
                                                                sPeriodoDocumentoProveedor = dtFechaInicioPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                            }

                                                            if (claseDocumento.TipoDocumentoChecklist == "Documento checklist")
                                                            {
                                                                if (bValidaDocumentoPorProyecto == true)
                                                                {
                                                                    foreach (var proyecto in searchResultsProyectoPorProveedor)
                                                                    {
                                                                        // Enviar la informacion del documento faltante a la BD
                                                                        oQuery.InsertarDocumentosFaltantesChecklist(
                                                                            bDelete,
                                                                            sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                            iProveedorID: organizacion.ObjVer.ID,
                                                                            sProyecto: proyecto.Title,
                                                                            iProyectoID: proyecto.ID,
                                                                            sCategoria: "Documento Faltante",
                                                                            sTipoDocumento: "Documento Proveedor",
                                                                            sNombreDocumento: szNombreClaseDocumento,
                                                                            iDocumentoID: 0,
                                                                            iClaseID: claseDocumento.DocumentoProveedor.ID,
                                                                            sVigencia: claseDocumento.VigenciaDocumentoProveedor,
                                                                            sPeriodo: sPeriodoDocumentoProveedor);
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    // Enviar la informacion del documento faltante a la BD
                                                                    oQuery.InsertarDocumentosFaltantesChecklist(
                                                                        bDelete,
                                                                        sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                        iProveedorID: organizacion.ObjVer.ID,
                                                                        sCategoria: "Documento Faltante",
                                                                        sTipoDocumento: "Documento Proveedor",
                                                                        sNombreDocumento: szNombreClaseDocumento,
                                                                        iDocumentoID: 0,
                                                                        iClaseID: claseDocumento.DocumentoProveedor.ID,
                                                                        sVigencia: claseDocumento.VigenciaDocumentoProveedor,
                                                                        sPeriodo: sPeriodoDocumentoProveedor);
                                                                }
                                                            }

                                                            // Crear nuevo periodo a partir de fecha fin que se convierte en la nueva fecha inicio
                                                            dtFechaInicioPeriodo = dtFechaFinPeriodo;
                                                            dtFechaFinPeriodo = ObtenerRangoDePeriodoDelDocumento
                                                            (
                                                                dtFechaInicioPeriodo,
                                                                claseDocumento.VigenciaDocumentoProveedor,
                                                                1
                                                            );

                                                            sFechaFin = "";
                                                            sFechaFin = dtFechaFinPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                                                            iDateCompare = DateTime.Compare
                                                            (
                                                                Convert.ToDateTime(sFechaActual), // t1
                                                                Convert.ToDateTime(sFechaFin)     // t2
                                                            );
                                                        }
                                                    }                                                                                                        

                                                    bNotification = true;
                                                    bConcatenateDocument = true;
                                                }
                                            }

                                            if (bActivaRelacionDeDocumentosVigentes)
                                            {
                                                // Relacion de documentos en la metadata del proveedor
                                                RelacionaDocumentosVigentes(
                                                    grupo.MasRecientesDocumentosRelacionados.ID,
                                                    organizacion,
                                                    oListaDocumentosVigentes);
                                            }

                                            SysUtils.ReportInfoToEventLog("Validacion de proyecto con contrato asociado en el proveedor: " + organizacion.Title);

                                            // Validar que proyecto asociado al proveedor tenga asociado un contrato (firmado, vigente), de lo contrario crear issue
                                            var ot_Proyecto = organizacion.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.Project");
                                            var cl_Contrato = organizacion.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Contrato");
                                            var cl_ProyectoServicioEspecializado = organizacion.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProyectoServicioEspecializado");
                                            var pd_Proveedor = organizacion.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
                                            var pd_ProyectosRelacionados = organizacion.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Project");
                                            var pd_EstatusProyectoServicioEspecializado = organizacion.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusProyectoServicioEspecializado");
                                            var pd_TipoContrato = organizacion.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.TipoDeContrato");
                                            var wf_CicloVidaContrato = organizacion.Vault.WorkflowOperations.GetWorkflowIDByAlias("MF.WF.ContractLifecycle");
                                            var wfs_Activo = organizacion.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("M-Files.CLM.State.ContractLifecycle.Active");

                                            var oLookupProveedor = new Lookup();
                                            var oLookupsProveedor = new Lookups();

                                            oLookupProveedor.Item = organizacion.ObjVer.ID;
                                            oLookupsProveedor.Add(-1, oLookupProveedor);

                                            // Buscar proyectos relacionados al proveedor
                                            var searchBuilderProyecto = new MFSearchBuilder(organizacion.Vault);
                                            searchBuilderProyecto.Deleted(false);
                                            searchBuilderProyecto.Class(cl_ProyectoServicioEspecializado);
                                            searchBuilderProyecto.Property(pd_Proveedor, MFDataType.MFDatatypeMultiSelectLookup, oLookupsProveedor);

                                            var searchResultsProyecto = searchBuilderProyecto.FindEx();

                                            if (searchResultsProyecto.Count > 0)
                                            {
                                                // Buscar contratos relacionados al proyecto
                                                var searchBuilderContrato = new MFSearchBuilder(organizacion.Vault);
                                                searchBuilderContrato.Deleted(false);
                                                searchBuilderContrato.Class(cl_Contrato);
                                                searchBuilderContrato.Property(pd_Proveedor, MFDataType.MFDatatypeMultiSelectLookup, oLookupsProveedor);
                                                searchBuilderContrato.Property(pd_TipoContrato, MFDataType.MFDatatypeLookup, 2);
                                                searchBuilderContrato.Property
                                                (
                                                    MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow,
                                                    MFDataType.MFDatatypeLookup,
                                                    wf_CicloVidaContrato
                                                );
                                                searchBuilderContrato.Property
                                                (
                                                    MFBuiltInPropertyDef.MFBuiltInPropertyDefState,
                                                    MFDataType.MFDatatypeLookup,
                                                    wfs_Activo
                                                );

                                                var searchResultsContrato = searchBuilderContrato.FindEx();

                                                foreach (var proyecto in searchResultsProyecto)
                                                {
                                                    SysUtils.ReportInfoToEventLog("Proyecto validado: " + proyecto.Title);

                                                    if (searchResultsContrato.Count > 0)
                                                    {
                                                        // Agregar propiedad de estatus de proyecto se y establecer el estatus: Con Contrato Asociado
                                                        var oLookup = new Lookup();
                                                        var oObjID = new ObjID();

                                                        oObjID.SetIDs
                                                        (
                                                            ObjType: ot_Proyecto,
                                                            ID: proyecto.ObjVer.ID
                                                        );

                                                        var checkedOutObjectVersion = proyecto.Vault.ObjectOperations.CheckOut(oObjID);

                                                        var oPropertyValue = new PropertyValue
                                                        {
                                                            PropertyDef = pd_EstatusProyectoServicioEspecializado
                                                        };

                                                        oLookup.Item = 5;

                                                        oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                                                        proyecto.Vault.ObjectPropertyOperations.SetProperty
                                                        (
                                                            ObjVer: checkedOutObjectVersion.ObjVer,
                                                            PropertyValue: oPropertyValue
                                                        );

                                                        proyecto.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);

                                                        // Relacionar proyecto y contrato ?
                                                        SysUtils.ReportInfoToEventLog("Proyecto con contrato asociado");
                                                    }
                                                    else
                                                    {
                                                        // Agregar propiedad de estatus de proyecto se y establecer el estatus: Sin Contrato Asociado
                                                        var oLookup = new Lookup();
                                                        var oObjID = new ObjID();

                                                        oObjID.SetIDs
                                                        (
                                                            ObjType: ot_Proyecto,
                                                            ID: proyecto.ObjVer.ID
                                                        );

                                                        var checkedOutObjectVersion = proyecto.Vault.ObjectOperations.CheckOut(oObjID);

                                                        var oPropertyValue = new PropertyValue
                                                        {
                                                            PropertyDef = pd_EstatusProyectoServicioEspecializado
                                                        };

                                                        oLookup.Item = 6;

                                                        oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                                                        proyecto.Vault.ObjectPropertyOperations.SetProperty
                                                        (
                                                            ObjVer: checkedOutObjectVersion.ObjVer,
                                                            PropertyValue: oPropertyValue
                                                        );

                                                        proyecto.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);

                                                        // Crear issue
                                                        CreateIssue(organizacion, proyecto);

                                                        SysUtils.ReportInfoToEventLog("Proyecto sin contrato asociado");
                                                    }
                                                }
                                            }

                                            // Busqueda Contacto Externo Servicio Especializado
                                            // Filtros: Nombre de proveedor, estatus
                                            var sbContactosExternosSE = new MFSearchBuilder(PermanentVault);
                                            sbContactosExternosSE.Deleted(false);
                                            sbContactosExternosSE.ObjType(ot_ContactoExternoSE);
                                            sbContactosExternosSE.Property
                                            (
                                                grupo.PropertyDefProveedorSEDocumentos, // Owner (Proveedor SE) - ID: 1730
                                                MFDataType.MFDatatypeMultiSelectLookup,
                                                oLookupsProveedor // organizacion.ObjVer.ID
                                            );
                                            sbContactosExternosSE.PropertyNot
                                            (
                                                pd_EstatusContactoExternoSE,
                                                MFDataType.MFDatatypeLookup,
                                                3 // Inactivo
                                            );

                                            var noContactos = sbContactosExternosSE.FindEx().Count;
                                            SysUtils.ReportInfoToEventLog("Contactos: " + noContactos);

                                            foreach (var contactoExterno in sbContactosExternosSE.FindEx())
                                            {
                                                bDelete = false;

                                                List<ObjVer> oListaDeDocumentosPorEmpleado = new List<ObjVer>();
                                                //bool bActivaActualizacionEstatusContactoExterno = false;
                                                var NombreContactoExterno = "";
                                                string sDocumentosContatenados = "";
                                                string sPeriodoSolicitadoDeDocumento = "";
                                                bool bActivaConcatenarContactoExterno = false;

                                                foreach (var claseEmpleadoLO in grupo.DocumentosEmpleado)
                                                {

                                                    if (claseEmpleadoLO.TipoValidacion == "Por empleado")
                                                    {
                                                        var sComparaFecha1 = "";
                                                        var sComparaFecha2 = "";
                                                        var dtFecha1 = new DateTime();
                                                        var dtFecha2 = new DateTime();
                                                        var dtFechaFinal = new DateTime();
                                                        var objVerDocumento1 = new ObjVer();
                                                        var objVerDocumento2 = new ObjVer();
                                                        var objVerDocumentoFinal = new ObjVer();
                                                        var vigenciaDocumentoCD = "";
                                                        bool bInicializarMetodoValidaVigenciaDocumentoCD = false;
                                                        string szNombreClaseDocumento = "";

                                                        //iNumeroMaxDocumentosLO = grupo.DocumentosProveedor.Count();

                                                        var searchBuilderDocumentosEmpleado = new MFSearchBuilder(PermanentVault);
                                                        searchBuilderDocumentosEmpleado.Deleted(false);
                                                        searchBuilderDocumentosEmpleado.Property
                                                        (
                                                            (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                                                            MFDataType.MFDatatypeLookup,
                                                            claseEmpleadoLO.DocumentoEmpleado.ID
                                                        );
                                                        searchBuilderDocumentosEmpleado.Property
                                                        (
                                                            grupo.EmpleadoContactoExterno,
                                                            MFDataType.MFDatatypeMultiSelectLookup,
                                                            contactoExterno.ObjVer.ID
                                                        );
                                                        searchBuilderDocumentosEmpleado.Property
                                                        (
                                                            grupo.PropertyDefProveedorSEDocumentos,
                                                            MFDataType.MFDatatypeMultiSelectLookup,
                                                            organizacion.ObjVer.ID
                                                        );
                                                        //searchBuilderDocumentosEmpleado.Property
                                                        //(
                                                        //    MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow,
                                                        //    MFDataType.MFDatatypeLookup,
                                                        //    grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID
                                                        //);
                                                        //searchBuilderDocumentosEmpleado.Property
                                                        //(
                                                        //    MFBuiltInPropertyDef.MFBuiltInPropertyDefState,
                                                        //    MFDataType.MFDatatypeLookup,
                                                        //    grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                        //);

                                                        if (searchBuilderDocumentosEmpleado.FindEx().Count > 0) // Se encontro al menos un documento
                                                        {
                                                            foreach (var documentoEmpleado in searchBuilderDocumentosEmpleado.FindEx())
                                                            {
                                                                oListaTodosLosDocumentosLO.Add(documentoEmpleado.ObjVer);

                                                                oPropertyValues = PermanentVault
                                                                    .ObjectPropertyOperations
                                                                    .GetProperties(documentoEmpleado.ObjVer);

                                                                ObjectClass oObjectClass = PermanentVault
                                                                    .ClassOperations
                                                                    .GetObjectClass(claseEmpleadoLO.DocumentoEmpleado.ID);

                                                                szNombreClaseDocumento = oObjectClass.Name;

                                                                if (oPropertyValues.IndexOf(grupo.Vigencia) != -1 &&
                                                                            !oPropertyValues.SearchForPropertyEx(grupo.Vigencia, true).TypedValue.IsNULL())
                                                                {
                                                                    if (oPropertyValues.IndexOf(grupo.FechaDeDocumento) != -1 &&
                                                                        !oPropertyValues.SearchForPropertyEx(grupo.FechaDeDocumento, true).TypedValue.IsNULL())
                                                                    {
                                                                        var fechaEmisionDocumentoCD = oPropertyValues
                                                                            .SearchForPropertyEx(grupo.FechaDeDocumento, true)
                                                                            .TypedValue
                                                                            .Value;

                                                                        // Comparar fecha de documentos (misma clase de documento) encontrados
                                                                        // Obtener el documento mas reciente y vigente
                                                                        if (sComparaFecha1 == "")
                                                                        {
                                                                            dtFecha1 = Convert.ToDateTime(fechaEmisionDocumentoCD);
                                                                            sComparaFecha1 = dtFecha1.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                            objVerDocumento1 = documentoEmpleado.ObjVer;
                                                                        }
                                                                        else
                                                                        {
                                                                            dtFecha2 = Convert.ToDateTime(fechaEmisionDocumentoCD);
                                                                            sComparaFecha2 = dtFecha2.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                            objVerDocumento2 = documentoEmpleado.ObjVer;
                                                                        }

                                                                        if (sComparaFecha1 != "" && sComparaFecha2 != "")
                                                                        {
                                                                            int iComparaFechasDeDocumentoChecklist = DateTime.Compare
                                                                            (
                                                                                Convert.ToDateTime(sComparaFecha1),
                                                                                Convert.ToDateTime(sComparaFecha2)
                                                                            );

                                                                            if (iComparaFechasDeDocumentoChecklist < 0)
                                                                            {
                                                                                sComparaFecha1 = "";
                                                                                objVerDocumentoFinal = objVerDocumento2;
                                                                                dtFechaFinal = dtFecha2;
                                                                            }
                                                                            else
                                                                            {
                                                                                sComparaFecha2 = "";
                                                                                objVerDocumentoFinal = objVerDocumento1;
                                                                                dtFechaFinal = dtFecha1;
                                                                            }

                                                                            bInicializarMetodoValidaVigenciaDocumentoCD = true;
                                                                        }

                                                                        // Si solo hay un documento, se establece el ID de objeto, la fecha de documento y la vigencia
                                                                        // directamente en el metodo ""
                                                                        if (searchBuilderDocumentosEmpleado.FindEx().Count == 1)
                                                                        {
                                                                            objVerDocumentoFinal = objVerDocumento1;
                                                                            dtFechaFinal = dtFecha1;
                                                                            bInicializarMetodoValidaVigenciaDocumentoCD = true;
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            if (bInicializarMetodoValidaVigenciaDocumentoCD)
                                                            {
                                                                // Obtener Vigencia de la clase verificada como la mas actual y vigente
                                                                var oPropertyValuesVigencia = new PropertyValues();

                                                                oPropertyValuesVigencia = PermanentVault
                                                                    .ObjectPropertyOperations
                                                                    .GetProperties(objVerDocumentoFinal);

                                                                vigenciaDocumentoCD = oPropertyValuesVigencia
                                                                    .SearchForPropertyEx(grupo.Vigencia.ID, true)
                                                                    .TypedValue
                                                                    .GetValueAsLocalizedText();

                                                                if (!(vigenciaDocumentoCD == "No Aplica")) // Valida vigencia de documento
                                                                {
                                                                    if (ValidarVigenciaDeDocumentoEnPeriodoActual(vigenciaDocumentoCD, dtFechaFinal) == false)
                                                                    {
                                                                        string ListaItems = LeerPlantilla(RutaLista);
                                                                        sChecklistDocumentName += ListaItems.Replace("[Documento]", szNombreClaseDocumento);

                                                                        if (oPropertyValues.IndexOf(pd_EstatusDocumento) != -1)
                                                                        {
                                                                            // Se modifica el estatus del documento a vencido
                                                                            ActualizarEstatusDocumento
                                                                            (
                                                                                "Documento",
                                                                                objVerDocumentoFinal,
                                                                                pd_EstatusDocumento,
                                                                                2,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoVencido.ID                                                                               
                                                                            );
                                                                        }

                                                                        // Obtener periodo del documento LO (mes de vencimiento)
                                                                        var PeriodoVencimientoDocumentoCD = ObtenerPeriodoDeVencimientoDelDocumento(dtFechaFinal, vigenciaDocumentoCD);

                                                                        if (PeriodoVencimientoDocumentoCD != "")
                                                                        {
                                                                            sPeriodoVencimientoDocumentoLO += PeriodoVencimientoDocumentoCD + "<br/>";
                                                                        }

                                                                        // Insertar informacion de documento vencido
                                                                        oQuery.InsertarDocumentosFaltantesChecklist(
                                                                            bDelete,
                                                                            sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                            iProveedorID: organizacion.ObjVer.ID,
                                                                            sEmpleado: contactoExterno.Title,
                                                                            iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                            sCategoria: "Documento Vencido",
                                                                            sTipoDocumento: "Documento Empleado",
                                                                            sNombreDocumento: szNombreClaseDocumento,
                                                                            iDocumentoID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                            iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                            sVigencia: vigenciaDocumentoCD,
                                                                            sPeriodo: PeriodoVencimientoDocumentoCD);

                                                                        bNotification = true;
                                                                        bConcatenateDocument = true;

                                                                        // Concatenar los documentos al contacto externo correspondiente
                                                                        bActivaConcatenarContactoExterno = true;
                                                                    }
                                                                    else
                                                                    {
                                                                        // Se agrega a la lista el documento vigente
                                                                        oListaDocumentosVigentes.Add(objVerDocumentoFinal);

                                                                        // Activa la relacion de objetos
                                                                        bActivaRelacionDeDocumentosVigentes = true;

                                                                        // Actualizar el estatus "Vigente" al documento
                                                                        ActualizarEstatusDocumento
                                                                        (
                                                                            "Documento",
                                                                            objVerDocumentoFinal,
                                                                            pd_EstatusDocumento,
                                                                            1,
                                                                            grupo.ConfigurationWorkflow.WorkflowDocumentoEmpleado.WorkflowValidacionesDocEmpleado.ID,
                                                                            grupo.ConfigurationWorkflow.WorkflowDocumentoEmpleado.EstadoDocumentoVigenteEmpleado.ID
                                                                        );
                                                                    }
                                                                }
                                                                else // Si el documento "No Aplica" validacion de vigencia
                                                                {
                                                                    oListaDocumentosVigentes.Add(objVerDocumentoFinal);
                                                                    bActivaRelacionDeDocumentosVigentes = true;

                                                                    // Actualizar el estatus "Vigente" al documento
                                                                    ActualizarEstatusDocumento
                                                                    (
                                                                        "Documento",
                                                                        objVerDocumentoFinal,
                                                                        pd_EstatusDocumento,
                                                                        1,
                                                                        grupo.ConfigurationWorkflow.WorkflowDocumentoEmpleado.WorkflowValidacionesDocEmpleado.ID,
                                                                        grupo.ConfigurationWorkflow.WorkflowDocumentoEmpleado.EstadoDocumentoVigenteEmpleado.ID
                                                                    );
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            szNombreClaseDocumento = claseEmpleadoLO.NombreClaseDocumento;

                                                            string ListaItemsEnviadosAEmpleado = LeerPlantilla(RutaListaEmpleado);
                                                            sDocumentosEnviadosAEmpleado += ListaItemsEnviadosAEmpleado.Replace("[DocumentoEmpleado]", szNombreClaseDocumento);
                                                            sPeriodosDeDocumentosEnviadosAEmpleado += "Faltante" + "<br/>";

                                                            if (bValidaDocumentoPorProyecto == true)
                                                            {
                                                                foreach (var proyecto in searchResultsProyectoPorProveedor)
                                                                {
                                                                    // Enviar la informacion del documento faltante a la BD
                                                                    oQuery.InsertarDocumentosFaltantesChecklist(
                                                                        bDelete,
                                                                        sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                        iProveedorID: organizacion.ObjVer.ID,
                                                                        sProyecto: proyecto.Title,
                                                                        iProyectoID: proyecto.ID,
                                                                        sEmpleado: contactoExterno.Title,
                                                                        iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                        sCategoria: "Documento Faltante",
                                                                        sTipoDocumento: "Documento Empleado",
                                                                        sNombreDocumento: szNombreClaseDocumento,
                                                                        iDocumentoID: 0,
                                                                        iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                        sPeriodo: "Faltante");
                                                                }
                                                            }
                                                            else
                                                            {
                                                                // Insertar informacion de documento vencido
                                                                oQuery.InsertarDocumentosFaltantesChecklist(
                                                                    bDelete,
                                                                    sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                    iProveedorID: organizacion.ObjVer.ID,
                                                                    sEmpleado: contactoExterno.Title,
                                                                    iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                    sCategoria: "Documento Faltante",
                                                                    sTipoDocumento: "Documento Empleado",
                                                                    sNombreDocumento: szNombreClaseDocumento,
                                                                    iDocumentoID: 0,
                                                                    iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                    sPeriodo: "Faltante");
                                                            }

                                                            // Concatenar los documentos al contacto externo correspondiente
                                                            bActivaConcatenarContactoExterno = true;
                                                        }
                                                    }

                                                    if (claseEmpleadoLO.TipoValidacion == "Por frecuencia de pago")
                                                    {
                                                        var sPeriodoDocumentoEmpleado = "";

                                                        oPropertyValues = PermanentVault
                                                            .ObjectPropertyOperations
                                                            .GetProperties(contactoExterno.ObjVer);

                                                        if (oPropertyValues.IndexOf(pd_FrecuenciaDePagoDeNomina) != -1)
                                                        {
                                                            if (!oPropertyValues.SearchForPropertyEx(pd_FrecuenciaDePagoDeNomina, true).TypedValue.IsNULL())
                                                            {
                                                                var frecuenciaPagoNomina = oPropertyValues
                                                                    .SearchForPropertyEx(pd_FrecuenciaDePagoDeNomina, true)
                                                                    .TypedValue
                                                                    .GetValueAsLocalizedText();

                                                                DateTime dtFechaInicioPeriodo = Convert.ToDateTime(fechaInicioProveedor);

                                                                DateTime dtFechaFinPeriodo = ObtenerRangoDePeriodoDelDocumento
                                                                (
                                                                    dtFechaInicioPeriodo,
                                                                    frecuenciaPagoNomina,
                                                                    1
                                                                );

                                                                // Validar fecha fin contra la fecha actual                                            
                                                                string sFechaFin = dtFechaFinPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                                                                int iDateCompare = DateTime.Compare
                                                                (
                                                                    Convert.ToDateTime(sFechaActual), // t1
                                                                    Convert.ToDateTime(sFechaFin)     // t2
                                                                );

                                                                while (iDateCompare >= 0)
                                                                {
                                                                    List<ObjVer> oDocumentosVigentesPorValidar = new List<ObjVer>();
                                                                    List<ObjVer> oDocumentosVencidos = new List<ObjVer>();

                                                                    // Busqueda de documentos del empleado
                                                                    var sbDocumentosEmpleadosLO = new MFSearchBuilder(PermanentVault);
                                                                    sbDocumentosEmpleadosLO.Deleted(false);
                                                                    sbDocumentosEmpleadosLO.Property
                                                                    (
                                                                        grupo.EmpleadoContactoExterno,
                                                                        MFDataType.MFDatatypeMultiSelectLookup,
                                                                        contactoExterno.ObjVer.ID
                                                                    );
                                                                    sbDocumentosEmpleadosLO.Property
                                                                    (
                                                                        (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                                                                        MFDataType.MFDatatypeLookup,
                                                                        claseEmpleadoLO.DocumentoEmpleado.ID
                                                                    );
                                                                    sbDocumentosEmpleadosLO.Property
                                                                    (
                                                                        grupo.PropertyDefProveedorSEDocumentos,
                                                                        MFDataType.MFDatatypeMultiSelectLookup,
                                                                        organizacion.ObjVer.ID
                                                                    );
                                                                    //sbDocumentosEmpleadosLO.Property
                                                                    //(
                                                                    //    MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow,
                                                                    //    MFDataType.MFDatatypeLookup,
                                                                    //    grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID
                                                                    //);
                                                                    //sbDocumentosEmpleadosLO.Property
                                                                    //(
                                                                    //    MFBuiltInPropertyDef.MFBuiltInPropertyDefState,
                                                                    //    MFDataType.MFDatatypeLookup,
                                                                    //    grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                    //);

                                                                    var documentosCount = sbDocumentosEmpleadosLO.FindEx().Count;
                                                                    SysUtils.ReportInfoToEventLog("Documentos: " + documentosCount);

                                                                    if (sbDocumentosEmpleadosLO.FindEx().Count > 0) // Se encontro al menos un documento
                                                                    {
                                                                        foreach (var documentoEmpleado in sbDocumentosEmpleadosLO.FindEx())
                                                                        {
                                                                            oPropertyValues = PermanentVault
                                                                                .ObjectPropertyOperations
                                                                                .GetProperties(documentoEmpleado.ObjVer);

                                                                            // Invocar fecha de pago de CFDI Nomina
                                                                            var pd_FechaDePago = PermanentVault
                                                                                .PropertyDefOperations
                                                                                .GetPropertyDefIDByAlias("CFDI.FechaPago.Texto");

                                                                            // Obtener fecha del documento
                                                                            var oFechaDeDocumento = oPropertyValues
                                                                                .SearchForPropertyEx(pd_FechaDePago, true)
                                                                                .TypedValue
                                                                                .Value;

                                                                            DateTime dtFechaDeDocumento = Convert.ToDateTime(oFechaDeDocumento);

                                                                            string sFechaDeDocumento = dtFechaDeDocumento.ToString("yyyy-MM-dd");

                                                                            DateTime? dtFechaFinVigencia = null;

                                                                            // Si existe la propiedad en la metadata del documento
                                                                            if (oPropertyValues.IndexOf(grupo.FechaFinVigencia) != -1)
                                                                            {
                                                                                if (!oPropertyValues.SearchForPropertyEx(grupo.FechaFinVigencia, true).TypedValue.IsNULL())
                                                                                {
                                                                                    // Obtener fecha fin de vigencia
                                                                                    var oFechaFinVigencia = oPropertyValues.SearchForPropertyEx(grupo.FechaFinVigencia, true).TypedValue.Value;
                                                                                    dtFechaFinVigencia = Convert.ToDateTime(oFechaFinVigencia);
                                                                                }
                                                                            }

                                                                            if (claseEmpleadoLO.TipoValidacionVigenciaDocumento == "Por periodo")
                                                                            {
                                                                                // Validar si la fecha del documento esta dentro del periodo obtenido
                                                                                if (ComparaFechaBaseContraUnaFechaInicioYFechaFin(
                                                                                sFechaDeDocumento,
                                                                                dtFechaInicioPeriodo,
                                                                                dtFechaFinPeriodo) == true)
                                                                                {
                                                                                    oDocumentosVigentesPorValidar.Add(documentoEmpleado.ObjVer);

                                                                                    // Actualizar el estatus "Vigente" al documento
                                                                                    ActualizarEstatusDocumento
                                                                                    (
                                                                                        "Documento",
                                                                                        documentoEmpleado.ObjVer,
                                                                                        pd_EstatusDocumento,
                                                                                        1,
                                                                                        grupo.ConfigurationWorkflow.WorkflowDocumentoEmpleado.WorkflowValidacionesDocEmpleado.ID,
                                                                                        grupo.ConfigurationWorkflow.WorkflowDocumentoEmpleado.EstadoDocumentoVigenteEmpleado.ID,
                                                                                        0,
                                                                                        documentoEmpleado,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                                    );
                                                                                }
                                                                                else
                                                                                {
                                                                                    oDocumentosVencidos.Add(documentoEmpleado.ObjVer);

                                                                                    // Agregar el estatus "Vencido" al documento
                                                                                    ActualizarEstatusDocumento
                                                                                    (
                                                                                        "Documento",
                                                                                        documentoEmpleado.ObjVer,
                                                                                        pd_EstatusDocumento,
                                                                                        2,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoVencido.ID,
                                                                                        0,
                                                                                        documentoEmpleado,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                                    );
                                                                                }
                                                                            }
                                                                            else // Es Por fecha de vigencia
                                                                            {
                                                                                // Validar si la fecha del documento esta dentro del periodo obtenido
                                                                                if (ComparaFechaBaseContraUnaFechaInicioYFechaFin(
                                                                                sFechaDeDocumento,
                                                                                dtFechaInicioPeriodo,
                                                                                dtFechaFinPeriodo) == true)
                                                                                {
                                                                                    oDocumentosVigentesPorValidar.Add(documentoEmpleado.ObjVer);
                                                                                }
                                                                                else
                                                                                {
                                                                                    oDocumentosVencidos.Add(documentoEmpleado.ObjVer);
                                                                                }

                                                                                // Validar vigencia tomando como referencia el ultimo periodo
                                                                                if (ValidarVigenciaDeDocumentoEnPeriodoActual(
                                                                                    frecuenciaPagoNomina,
                                                                                    dtFechaDeDocumento,
                                                                                    dtFechaFinVigencia) == true)
                                                                                {
                                                                                    // Actualizar el estatus "Vigente" al documento
                                                                                    ActualizarEstatusDocumento
                                                                                    (
                                                                                        "Documento",
                                                                                        documentoEmpleado.ObjVer,
                                                                                        pd_EstatusDocumento,
                                                                                        1,
                                                                                        grupo.ConfigurationWorkflow.WorkflowDocumentoEmpleado.WorkflowValidacionesDocEmpleado.ID,
                                                                                        grupo.ConfigurationWorkflow.WorkflowDocumentoEmpleado.EstadoDocumentoVigenteEmpleado.ID,
                                                                                        0,
                                                                                        documentoEmpleado,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                                    );
                                                                                }
                                                                                else
                                                                                {
                                                                                    // Agregar el estatus "Vencido" al documento
                                                                                    ActualizarEstatusDocumento
                                                                                    (
                                                                                        "Documento",
                                                                                        documentoEmpleado.ObjVer,
                                                                                        pd_EstatusDocumento,
                                                                                        2,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoVencido.ID,
                                                                                        0,
                                                                                        documentoEmpleado,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.WorkflowValidacionesChecklist.ID,
                                                                                        grupo.ConfigurationWorkflow.WorkflowChecklist.EstadoDocumentoProcesado.ID
                                                                                    );
                                                                                }
                                                                            }                                                                            
                                                                        }

                                                                        // Fecha inicio y fin del periodo validado
                                                                        var sPeriodoDocumentoFaltante = ObtenerPeriodoDeDocumentoFaltante
                                                                        (
                                                                            frecuenciaPagoNomina,
                                                                            dtFechaInicioPeriodo,
                                                                            dtFechaFinPeriodo
                                                                        );

                                                                        if (oDocumentosVigentesPorValidar.Count > 0)
                                                                        {
                                                                            // Agregar a la lista, el documento encontrado
                                                                            oListaDeDocumentosPorEmpleado.Add(oDocumentosVigentesPorValidar[0]);

                                                                            // Actualizar el estatus del Contacto Externo
                                                                            ActualizarEstatusDocumento
                                                                            (
                                                                                "Contacto Externo",
                                                                                contactoExterno.ObjVer,
                                                                                pd_EstatusContactoExternoSE,
                                                                                2, 0, 0,
                                                                                ot_ContactoExternoSE
                                                                            );
                                                                        }
                                                                        else
                                                                        {
                                                                            // Se agrega al correo el nombre del documento no encontrado
                                                                            // en el periodo validado
                                                                            string ListaItemsEmpleado = LeerPlantilla(RutaListaEmpleado);
                                                                            sDocumentosContatenados += ListaItemsEmpleado.Replace("[DocumentoEmpleado]", claseEmpleadoLO.NombreClaseDocumento);

                                                                            sPeriodoDocumentoEmpleado = sPeriodoDocumentoFaltante;

                                                                            sPeriodoSolicitadoDeDocumento += sPeriodoDocumentoEmpleado + "<br/>";

                                                                            // Modificar formato de la fecha del periodo
                                                                            if (frecuenciaPagoNomina == "Mensual")
                                                                            {
                                                                                sPeriodoDocumentoEmpleado = dtFechaInicioPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                            }

                                                                            // Insert
                                                                            oQuery.InsertarDocumentosFaltantesChecklist(
                                                                                bDelete,
                                                                                sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                                iProveedorID: organizacion.ObjVer.ID,
                                                                                sEmpleado: contactoExterno.Title,
                                                                                iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                                sCategoria: "Documento Vencido",
                                                                                sTipoDocumento: "Documento Empleado",
                                                                                sNombreDocumento: claseEmpleadoLO.NombreClaseDocumento,
                                                                                iDocumentoID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                                iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                                sVigencia: frecuenciaPagoNomina,
                                                                                sPeriodo: sPeriodoDocumentoEmpleado);

                                                                            // Modificar el estatus del Contacto Externo
                                                                            ActualizarEstatusDocumento
                                                                            (
                                                                                "Contacto Externo",
                                                                                contactoExterno.ObjVer,
                                                                                pd_EstatusContactoExternoSE,
                                                                                1, 0, 0,
                                                                                ot_ContactoExternoSE
                                                                            );

                                                                            // Concatenar los documentos al contacto externo correspondiente
                                                                            bActivaConcatenarContactoExterno = true;
                                                                        }

                                                                        // Enviar los documentos vencidos a la tabla de documentos faltantes checklist 
                                                                        foreach (var documento in oDocumentosVencidos)
                                                                        {
                                                                            sPeriodoDocumentoEmpleado = sPeriodoDocumentoFaltante;

                                                                            // Modificar formato de la fecha del periodo
                                                                            if (frecuenciaPagoNomina == "Mensual")
                                                                            {
                                                                                sPeriodoDocumentoEmpleado = dtFechaInicioPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                            }

                                                                            SysUtils.ReportInfoToEventLog("Insertando documento: " + documento.ID + " en la tabla DocumentosCaducados");

                                                                            if (bValidaDocumentoPorProyecto == true)
                                                                            {
                                                                                foreach (var proyecto in searchResultsProyectoPorProveedor)
                                                                                {
                                                                                    var pd_DocumentosRelacionadosAlProyecto = proyecto.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Document");
                                                                                    var oPropertiesProyecto = proyecto.Properties;

                                                                                    var oListDocumentosProyecto = oPropertiesProyecto
                                                                                        .SearchForPropertyEx(pd_DocumentosRelacionadosAlProyecto, true)
                                                                                        .TypedValue
                                                                                        .GetValueAsLookups()
                                                                                        .ToObjVerExs(proyecto.Vault);

                                                                                    foreach (var documentoProyecto in oListDocumentosProyecto)
                                                                                    {
                                                                                        if (documento.ID == documentoProyecto.ObjVer.ID)
                                                                                        {
                                                                                            // Insert
                                                                                            oQuery.InsertarDocumentosCaducados(
                                                                                                sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                                                iProveedorID: organizacion.ObjVer.ID,
                                                                                                sProyecto: proyecto.Title,
                                                                                                iProyectoID: proyecto.ID,
                                                                                                sEmpleado: contactoExterno.Title,
                                                                                                iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                                                sCategoria: "Documento Vencido",
                                                                                                sTipoDocumento: "Documento Empleado",
                                                                                                sNombreDocumento: claseEmpleadoLO.NombreClaseDocumento,
                                                                                                iDocumentoID: documento.ID,
                                                                                                iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                                                sVigencia: frecuenciaPagoNomina,
                                                                                                sPeriodo: sPeriodoDocumentoEmpleado);
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            // Insert
                                                                                            oQuery.InsertarDocumentosCaducados(
                                                                                                sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                                                iProveedorID: organizacion.ObjVer.ID,
                                                                                                sProyecto: proyecto.Title,
                                                                                                iProyectoID: proyecto.ID,
                                                                                                sEmpleado: contactoExterno.Title,
                                                                                                iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                                                sCategoria: "Documento Faltante",
                                                                                                sTipoDocumento: "Documento Empleado",
                                                                                                sNombreDocumento: claseEmpleadoLO.NombreClaseDocumento,
                                                                                                iDocumentoID: 0,
                                                                                                iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                                                sVigencia: frecuenciaPagoNomina,
                                                                                                sPeriodo: sPeriodoDocumentoEmpleado);
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                // Insert
                                                                                oQuery.InsertarDocumentosCaducados(
                                                                                    sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                                    iProveedorID: organizacion.ObjVer.ID,
                                                                                    sEmpleado: contactoExterno.Title,
                                                                                    iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                                    sCategoria: "Documento Vencido",
                                                                                    sTipoDocumento: "Documento Empleado",
                                                                                    sNombreDocumento: claseEmpleadoLO.NombreClaseDocumento,
                                                                                    iDocumentoID: documento.ID,
                                                                                    iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                                    sVigencia: frecuenciaPagoNomina,
                                                                                    sPeriodo: sPeriodoDocumentoEmpleado);
                                                                            }
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        // Si no se encuentra ningun documento de la clase de documento buscada
                                                                        // Se agrega la clase de documento en el correo para que se suba a la boveda
                                                                        string ListaItemsEmpleado = LeerPlantilla(RutaListaEmpleado);
                                                                        sDocumentosContatenados += ListaItemsEmpleado.Replace("[DocumentoEmpleado]", claseEmpleadoLO.NombreClaseDocumento);

                                                                        // Dar formato a periodo faltante de documento del proveedor
                                                                        var sPeriodoDocumentoFaltante = ObtenerPeriodoDeDocumentoFaltante
                                                                        (
                                                                            frecuenciaPagoNomina,
                                                                            dtFechaInicioPeriodo,
                                                                            dtFechaFinPeriodo
                                                                        );

                                                                        sPeriodoDocumentoEmpleado = sPeriodoDocumentoFaltante; //sMesAnioFechaInicioPeriodo;

                                                                        sPeriodoSolicitadoDeDocumento += sPeriodoDocumentoEmpleado + "<br/>";

                                                                        // Modificar formato de la fecha del periodo para enviarla a la BD
                                                                        if (frecuenciaPagoNomina == "Mensual")
                                                                        {
                                                                            sPeriodoDocumentoEmpleado = dtFechaInicioPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                        }

                                                                        if (bValidaDocumentoPorProyecto == true)
                                                                        {
                                                                            foreach (var proyecto in searchResultsProyectoPorProveedor)
                                                                            {
                                                                                // Enviar la informacion del documento faltante a la BD
                                                                                oQuery.InsertarDocumentosFaltantesChecklist(
                                                                                    bDelete,
                                                                                    sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                                    iProveedorID: organizacion.ObjVer.ID,
                                                                                    sProyecto: proyecto.Title,
                                                                                    iProyectoID: proyecto.ID,
                                                                                    sEmpleado: contactoExterno.Title,
                                                                                    iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                                    sCategoria: "Documento Faltante",
                                                                                    sTipoDocumento: "Documento Empleado",
                                                                                    sNombreDocumento: claseEmpleadoLO.NombreClaseDocumento,
                                                                                    iDocumentoID: 0,
                                                                                    iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                                    sVigencia: frecuenciaPagoNomina,
                                                                                    sPeriodo: sPeriodoDocumentoEmpleado);
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            // Enviar la informacion del documento faltante a la BD
                                                                            oQuery.InsertarDocumentosFaltantesChecklist(
                                                                                bDelete,
                                                                                sProveedor: nombreOTituloObjetoPadre.ToString(),
                                                                                iProveedorID: organizacion.ObjVer.ID,
                                                                                sEmpleado: contactoExterno.Title,
                                                                                iEmpleadoID: contactoExterno.ObjVer.ID,
                                                                                sCategoria: "Documento Faltante",
                                                                                sTipoDocumento: "Documento Empleado",
                                                                                sNombreDocumento: claseEmpleadoLO.NombreClaseDocumento,
                                                                                iDocumentoID: 0,
                                                                                iClaseID: claseEmpleadoLO.DocumentoEmpleado.ID,
                                                                                sVigencia: frecuenciaPagoNomina,
                                                                                sPeriodo: sPeriodoDocumentoEmpleado);
                                                                        }

                                                                        // Modificar el estatus del Contacto Externo
                                                                        ActualizarEstatusDocumento
                                                                        (
                                                                            "Contacto Externo",
                                                                            contactoExterno.ObjVer,
                                                                            pd_EstatusContactoExternoSE,
                                                                            1, 0, 0,
                                                                            ot_ContactoExternoSE
                                                                        );

                                                                        // Concatenar los documentos al contacto externo correspondiente
                                                                        bActivaConcatenarContactoExterno = true;
                                                                    }

                                                                    // Crear nuevo periodo a partir de fecha fin que se convierte en la nueva fecha inicio
                                                                    dtFechaInicioPeriodo = dtFechaFinPeriodo;
                                                                    dtFechaFinPeriodo = ObtenerRangoDePeriodoDelDocumento
                                                                    (
                                                                        dtFechaInicioPeriodo,
                                                                        frecuenciaPagoNomina,
                                                                        1
                                                                    );

                                                                    sFechaFin = "";
                                                                    sFechaFin = dtFechaFinPeriodo.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                                                                    iDateCompare = DateTime.Compare
                                                                    (
                                                                        Convert.ToDateTime(sFechaActual), // t1
                                                                        Convert.ToDateTime(sFechaFin)     // t2
                                                                    );
                                                                }

                                                                // Relacionar documentos en Contacto Externo
                                                                if (oListaDeDocumentosPorEmpleado.Count > 0)
                                                                {
                                                                    //RelacionaDocumentosVigentes(
                                                                    //    pd_DocumentosSEContactoExterno,
                                                                    //    contactoExterno,
                                                                    //    oListaDeDocumentosPorEmpleado);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                                // Concatenar nombre de contacto externo y los documentos no encontrados
                                                int valorEmpleado = 1;
                                                if (bActivaConcatenarContactoExterno)
                                                {
                                                    NombreContactoExterno = contactoExterno.Title;
                                                    string TbodyDocumentosEmpleado = LeerPlantilla(RutaTbodyEmpleado);
                                                    string TbodyRowColor = "";

                                                    if ((valorEmpleado % 2) == 0)
                                                    {
                                                        TbodyRowColor = "#eee";
                                                    }
                                                    else
                                                    {
                                                        TbodyRowColor = "#ffff";
                                                    }

                                                    sDocumentosContatenados += sDocumentosEnviadosAEmpleado;
                                                    sPeriodoSolicitadoDeDocumento += sPeriodosDeDocumentosEnviadosAEmpleado;

                                                    sBodyMessageDocumentsEmp = TbodyDocumentosEmpleado.Replace("[ListaEmpleado]", sDocumentosContatenados);
                                                    sBodyMessageDocumentsEmp = sBodyMessageDocumentsEmp.Replace("[Empleado]", NombreContactoExterno);
                                                    sBodyMessageDocumentsEmp = sBodyMessageDocumentsEmp.Replace("[trColor]", TbodyRowColor);
                                                    sBodyMessageDocumentsEmp = sBodyMessageDocumentsEmp.Replace("[Mes]", sPeriodoSolicitadoDeDocumento); //"PENDIENTE"

                                                    tBodyEmpleado += sBodyMessageDocumentsEmp;

                                                    // Limpiar documentos concatenados para evitar duplicidad
                                                    sDocumentosEnviadosAEmpleado = "";
                                                    sPeriodosDeDocumentosEnviadosAEmpleado = "";
                                                }
                                            }

                                            int valor = 1;

                                            if (bConcatenateDocument == true)
                                            {
                                                string TbodyDocumentos = LeerPlantilla(RutaTbody);
                                                string TbodyRowColor = "";

                                                if ((valor % 2) == 0)
                                                {
                                                    TbodyRowColor = "#eee";
                                                }
                                                else
                                                {
                                                    TbodyRowColor = "#ffff";
                                                }

                                                sBodyMessageDocuments = TbodyDocumentos.Replace("[Lista]", sChecklistDocumentName);
                                                sBodyMessageDocuments = sBodyMessageDocuments.Replace("[Documento]", "Documentos por actualizar");
                                                sBodyMessageDocuments = sBodyMessageDocuments.Replace("[trColor]", TbodyRowColor);
                                                sBodyMessageDocuments = sBodyMessageDocuments.Replace("[Mes]", sPeriodoVencimientoDocumentoLO);

                                                tBody += sBodyMessageDocuments;
                                            }
                                        }

                                        // Enviar notificacion (correo)
                                        if (bNotification)
                                        {
                                            // Agregar estatus "Documentos Pendientes" en el proveedor
                                            ActualizarEstatusDocumento
                                            (
                                                "Proveedor",
                                                organizacion.ObjVer,
                                                pd_EstatusProveedor,
                                                3, 0, 0,
                                                grupo.ValidacionOrganizacion.ObjetoOrganizacion
                                            );

                                            if (contactosAdministradores.Count > 0)
                                            {
                                                List<string> sEmails = new List<string>();

                                                // Extraer el email de contactos externos en proveedor
                                                foreach (var contacto in contactosAdministradores)
                                                {
                                                    oPropertyValues = PermanentVault
                                                        .ObjectPropertyOperations
                                                        .GetProperties(contacto.ObjVer);

                                                    var emailContactoExterno = oPropertyValues
                                                        .SearchForProperty(Configuration.ConfigurationServiciosGenerales.ConfigurationNotificaciones.PDEmail)
                                                        .TypedValue
                                                        .Value;

                                                    sEmails.Add(emailContactoExterno.ToString());
                                                }

                                                // Creacion de mensaje del correo
                                                Plantilla = Plantilla.Replace("[Proveedor]", nombreOTituloObjetoPadre.ToString());
                                                Plantilla = Plantilla.Replace("[Tbody]", tBody);
                                                Plantilla = Plantilla.Replace("[TbodyEmpleado]", tBodyEmpleado);
                                                Plantilla = Plantilla.Replace("[Cliente]", nombreOTituloObjetoPadre.ToString());

                                                sMainBodyMessage = Plantilla;

                                                LinkedResource ImgBanner = new LinkedResource(RutaBanner, MediaTypeNames.Image.Jpeg);
                                                LinkedResource ImgCloud = new LinkedResource(RutaCloud, MediaTypeNames.Image.Jpeg);
                                                LinkedResource ImgFooter = new LinkedResource(RutaFooter, MediaTypeNames.Image.Jpeg);

                                                ImgBanner.ContentId = "ImgBanner";
                                                ImgCloud.ContentId = "ImgCloud";
                                                ImgFooter.ContentId = "ImgFooter";

                                                AlternateView AV = AlternateView.CreateAlternateViewFromString(sMainBodyMessage, null, MediaTypeNames.Text.Html);
                                                AV.LinkedResources.Add(ImgBanner);
                                                AV.LinkedResources.Add(ImgCloud);
                                                AV.LinkedResources.Add(ImgFooter);

                                                Email oEmail = new Email();

                                                if (oEmail.Enviar(AV, sEmails,
                                                    Configuration.ConfigurationServiciosGenerales.ConfigurationNotificaciones.EmailService,
                                                    Configuration.ConfigurationServiciosGenerales.ConfigurationNotificaciones.HostService,
                                                    Configuration.ConfigurationServiciosGenerales.ConfigurationNotificaciones.PortService,
                                                    Configuration.ConfigurationServiciosGenerales.ConfigurationNotificaciones.UsernameService,
                                                    Configuration.ConfigurationServiciosGenerales.ConfigurationNotificaciones.PasswordService) == true)
                                                {
                                                    SysUtils.ReportInfoToEventLog("Fin del proceso, el correo ha sido enviado exitosamente.");
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // Si no se activa la notificacion para el proveedor es porque tiene todos sus documentos al dia
                                            // Agregar estatus "Proveedor Actualizado" en el proveedor
                                            ActualizarEstatusDocumento
                                            (
                                                "Proveedor",
                                                organizacion.ObjVer,
                                                pd_EstatusProveedor,
                                                4, 0, 0,
                                                grupo.ValidacionOrganizacion.ObjetoOrganizacion
                                            );
                                        }
                                    }
                                }
                            }
                        }
                    }                    
                });                
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Ocurrio un error. " + ex.Message);
            }
        }

        private List<ObjVerEx> GetExistingIssues(ObjVerEx envProveedor)
        {
            var ot_Issue = envProveedor.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.Issue");
            var cl_IssueServiciosEspecializados = envProveedor.Vault.ClassOperations.GetObjectClassIDByAlias("CL.IssueServiciosEspecializados");

            // Busqueda de issues existentes en la boveda
            var searchBuilder = new MFSearchBuilder(envProveedor.Vault);
            searchBuilder.Deleted(false); // No eliminados
            searchBuilder.ObjType(ot_Issue);
            searchBuilder.Class(cl_IssueServiciosEspecializados);

            var results = searchBuilder.FindEx();

            return results;
        }

        private void CreateIssue(ObjVerEx oProveedor, ObjVerEx oProyecto)
        {
            var ot_Issue = oProveedor.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.Issue");
            var cl_IssueServiciosEspecializados = oProveedor.Vault.ClassOperations.GetObjectClassIDByAlias("CL.IssueServiciosEspecializados");
            var pd_Proveedor = oProveedor.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_Proyecto = oProveedor.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Projects");
            var pd_Descripcion = oProveedor.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("M-Files.CLM.Property.Description");
            var pd_TipoIncidencia = oProveedor.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.IssueType");
            var pd_Severidad = oProveedor.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Severity");
            var wf_IssueProcessing = oProveedor.Vault.WorkflowOperations.GetWorkflowIDByAlias("MF.WF.IssueProcessing");
            var wfs_Submitted = oProveedor.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("M-Files.CLM.State.IssueProcessing.Submitted");

            var oLookupsIssueType = new Lookups();
            var oLookupsProveedor = new Lookups();
            var oLookupsProyecto = new Lookups();
            var oLookupIssueType = new Lookup();
            var oLookupProveedor = new Lookup();
            var oLookupProyecto = new Lookup();

            oLookupIssueType.Item = 1;
            oLookupsIssueType.Add(-1, oLookupIssueType);

            oLookupProveedor.Item = oProveedor.ObjVer.ID;
            oLookupsProveedor.Add(-1, oLookupProveedor);

            oLookupProyecto.Item = oProyecto.ObjVer.ID;
            oLookupsProyecto.Add(-1, oLookupProyecto);

            // Antes de crear el issue validar que aun no exista uno ya creado para el proveedor y proyecto
            var searchBuilderIssue = new MFSearchBuilder(oProveedor.Vault);
            searchBuilderIssue.Deleted(false); // No eliminados
            searchBuilderIssue.ObjType(ot_Issue);
            searchBuilderIssue.Property(pd_Proveedor, MFDataType.MFDatatypeMultiSelectLookup, oLookupsProveedor);
            searchBuilderIssue.Property(pd_Proyecto, MFDataType.MFDatatypeMultiSelectLookup, oLookupsProyecto);
            var searchResultsIssue = searchBuilderIssue.FindEx();

            if (searchResultsIssue.Count == 0)
            {
                // Generar el numero consecutivo del siguiente issue a crear
                var issues = GetExistingIssues(oProveedor);
                var issuesCount = issues.Count;
                var noConsecutivo = issuesCount + 1;
                var sNombreOTitulo = "Issue #" + noConsecutivo;

                var createBuilderIssue = new MFPropertyValuesBuilder(oProveedor.Vault);
                createBuilderIssue.SetClass(cl_IssueServiciosEspecializados);
                createBuilderIssue.Add
                (
                    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle,
                    MFDataType.MFDatatypeText,
                    sNombreOTitulo // Name or title
                );
                createBuilderIssue.Add(pd_Proveedor, MFDataType.MFDatatypeMultiSelectLookup, oLookupsProveedor);
                createBuilderIssue.Add(pd_Proyecto, MFDataType.MFDatatypeMultiSelectLookup, oLookupProyecto);
                createBuilderIssue.Add(pd_Descripcion, MFDataType.MFDatatypeMultiLineText, "El proyecto no contiene un contrato relacionado");
                createBuilderIssue.Add(pd_TipoIncidencia, MFDataType.MFDatatypeMultiSelectLookup, oLookupsIssueType);
                createBuilderIssue.Add(pd_Severidad, MFDataType.MFDatatypeLookup, 2);
                createBuilderIssue.SetWorkflowState(wf_IssueProcessing, wfs_Submitted);

                // Tipo de objeto a crear
                var objectTypeId = ot_Issue;

                // Finaliza la creacion del issue
                var objectVersion = oProveedor.Vault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeId,
                    createBuilderIssue.Values,
                    CheckIn: true
                );
            }            
        }

        private bool BuscaDocumentosChecklistFaltantes(
            int iPropertyDef,
            int iPropertyDefChecklist,
            int ObjetoPadreID,
            int ObjChecklistID,
            int iPropertyDefFechaEmisionChecklist,
            DateTime dtFechaEmisionDocumento,
            string sVigenciaDocumentChecklist = "",
            int iPropDefVigenciaClaseChecklist = 0)
        {
            bool bExisteDocumento = false;
            var resultadosDeBusqueda = new List<ObjVerEx>();
            var oPropertyValues = new PropertyValues();

            // Inicializar objeto de busqueda
            var searchBuilder = new MFSearchBuilder(this.PermanentVault);

            // Filtros de busqueda
            searchBuilder.Deleted(false);
            searchBuilder.Property
            (
                iPropertyDef,
                MFDataType.MFDatatypeMultiSelectLookup,
                ObjetoPadreID
            );
            searchBuilder.Property
            (
                iPropertyDefChecklist,
                MFDataType.MFDatatypeLookup,
                ObjChecklistID
            );

            resultadosDeBusqueda = searchBuilder.FindEx();

            if (resultadosDeBusqueda.Count > 0)
            {
                // Validar Fecha Emision de Documento contra Fecha Emision de Documento checklist
                ObjVerEx documentoChecklist = resultadosDeBusqueda[0];

                oPropertyValues = PermanentVault
                    .ObjectPropertyOperations
                    .GetProperties(documentoChecklist.ObjVer);

                if (sVigenciaDocumentChecklist == "")
                {
                    // Obtener vigencia de la clase verificada
                    sVigenciaDocumentChecklist = oPropertyValues
                        .SearchForPropertyEx(iPropDefVigenciaClaseChecklist, true)
                        .TypedValue
                        .GetValueAsLocalizedText();
                }

                var fechaEmisionChecklist = oPropertyValues
                    .SearchForProperty(iPropertyDefFechaEmisionChecklist)
                    .TypedValue
                    .Value;

                DateTime dtFechaEmisionTipoDocumento = Convert.ToDateTime(fechaEmisionChecklist);

                switch (sVigenciaDocumentChecklist)
                {
                    case "Semanal":

                        if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist(
                            dtFechaEmisionDocumento,
                            dtFechaEmisionTipoDocumento) <= 7)
                        {
                            bExisteDocumento = true;
                        }

                        break;

                    case "Mensual":

                        if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist(
                            dtFechaEmisionDocumento,
                            dtFechaEmisionTipoDocumento) <= 30)
                        {
                            bExisteDocumento = true;
                        }

                        break;

                    case "Bimestral":

                        if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist(
                            dtFechaEmisionDocumento,
                            dtFechaEmisionTipoDocumento) <= 60)
                        {
                            bExisteDocumento = true;
                        }

                        break;

                    case "Trimestral":

                        if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist(
                            dtFechaEmisionDocumento,
                            dtFechaEmisionTipoDocumento) <= 90)
                        {
                            bExisteDocumento = true;
                        }

                        break;

                    case "Anual":

                        if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist(
                            dtFechaEmisionDocumento,
                            dtFechaEmisionTipoDocumento) <= 365)
                        {
                            bExisteDocumento = true;
                        }

                        break;
                }
            }

            return bExisteDocumento;
        }

        private bool ValidarVigenciaDeDocumentoEnPeriodoActual(string sVigenciaDocumento, DateTime dtFechaDocumento, DateTime? dtFechaFinVigencia = null)
        {
            bool bDocumentoVigente = false;
            
            DateTime dtFechaFin;

            if (dtFechaFinVigencia != null)
            {
                dtFechaFin = dtFechaFinVigencia.Value;
            }
            else
            {
                dtFechaFin = DateTime.Today;
            }

            DateTime dtFechaInicio = ObtenerRangoDePeriodoDelDocumento(dtFechaFin, sVigenciaDocumento, 0);

            string sFechaDocumento = dtFechaDocumento.ToString("yyyy-MM-dd");

            if (ComparaFechaBaseContraUnaFechaInicioYFechaFin(sFechaDocumento, dtFechaInicio, dtFechaFin) == true)
            {
                bDocumentoVigente = true;
            }
         
            return bDocumentoVigente;
        }

        private int CalculaDiasEntreDocumentoPadreYDocumentosChecklist(
            DateTime dtFechaEmisionDocumento,
            DateTime dtFechaEmisionChecklist)
        {
            TimeSpan diferenciaDeFechas = new TimeSpan();
            int diferenciaEnDias = 0;

            string sFechaUno = dtFechaEmisionDocumento.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            string sFechaDos = dtFechaEmisionChecklist.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            int iComparaFechasDeEmision = DateTime.Compare
            (
                Convert.ToDateTime(sFechaUno),
                Convert.ToDateTime(sFechaDos)
            );

            if (iComparaFechasDeEmision < 0)
            {
                diferenciaDeFechas = dtFechaEmisionChecklist - dtFechaEmisionDocumento;
            }
            else if (iComparaFechasDeEmision > 0)
            {
                diferenciaDeFechas = dtFechaEmisionDocumento - dtFechaEmisionChecklist;
            }
            else
            {
                diferenciaDeFechas = dtFechaEmisionChecklist - dtFechaEmisionDocumento;
            }

            diferenciaEnDias = diferenciaDeFechas.Days;

            return diferenciaEnDias;
        }

        private bool ComparaFechaBaseContraUnaFechaInicioYFechaFin(
            string sFechaBase,
            DateTime? dtFechaInicio,
            DateTime? dtFechaFin)
        {
            bool resultado = false; // El resultado se inicializa en false

            DateTime _dtFechaInicio = Convert.ToDateTime(dtFechaInicio);
            string sFechaInicio = _dtFechaInicio.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            DateTime _dtFechaFin = Convert.ToDateTime(dtFechaFin);
            string sFechaFin = _dtFechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            int ComparaFechaActualConFechaInicial = DateTime.Compare
            (
                Convert.ToDateTime(sFechaBase),  // t1
                Convert.ToDateTime(sFechaInicio) // t2
            );

            int ComparaFechaActualConFechaFin = DateTime.Compare
            (
                Convert.ToDateTime(sFechaBase), // t1
                Convert.ToDateTime(sFechaFin)   // t2
            );

            if (ComparaFechaActualConFechaInicial >= 0 && ComparaFechaActualConFechaFin <= 0)
            {
                // Si se cumplen ambas condiciones Fecha Actual esta dentro del rango
                resultado = true;
            }

            return resultado;
        }

        private bool ValidaFechaDocumentoEmpleadoConUltimoMes(
            DateTime? dtFechaInicio,
            DateTime? dtFechaFin,
            string sFechaDocumentoEmpleado)
        {
            bool resultado = false; // El resultado se inicializa en false

            DateTime _dtFechaInicio = Convert.ToDateTime(dtFechaInicio);
            string sFechaInicio = _dtFechaInicio.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            DateTime _dtFechaFin = Convert.ToDateTime(dtFechaFin);
            string sFechaFin = _dtFechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            int ComparaFechaActualConFechaInicial = DateTime.Compare
            (
                Convert.ToDateTime(sFechaDocumentoEmpleado),
                Convert.ToDateTime(sFechaInicio)
            );

            int ComparaFechaActualConFechaFin = DateTime.Compare
            (
                Convert.ToDateTime(sFechaDocumentoEmpleado),
                Convert.ToDateTime(sFechaFin)
            );

            if (ComparaFechaActualConFechaInicial >= 0 && ComparaFechaActualConFechaFin <= 0)
            {
                // Si se cumplen ambas condiciones Fecha Actual esta dentro del rango
                resultado = true;
            }

            return resultado;
        }

        private void ActualizarWorkflowValidacionManual(ObjVerEx oObjVerEx, int iWorkflow, int iState)
        {
            var oWorkflowstate = new ObjectVersionWorkflowState();
            var oObjID = new ObjID();

            oObjID.SetIDs
            (
                ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                ID: oObjVerEx.ObjVer.ID
            );

            ObjVer checkedOutObjectVersion = oObjVerEx.Vault.ObjectOperations.GetLatestObjVerEx(oObjID, true);
            checkedOutObjectVersion = oObjVerEx.Vault.ObjectOperations.CheckOut(oObjID).ObjVer;

            oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iWorkflow);
            oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iState);
            oObjVerEx.Vault.ObjectPropertyOperations.SetWorkflowStateEx(checkedOutObjectVersion, oWorkflowstate);

            oObjVerEx.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion);
        }

        private void ActualizarEstatusDocumento(string sTipoObjeto, ObjVer oObjVer, int iPropertyDefEstatus, int iEstatusIdValue, int iWorkflow, int iState, int iObjectTypeNoDocumento = 0, ObjVerEx oObjVerEx = null, int iWorkflowValidacionesChecklist = 0, int iStateDocumentoProcesado = 0)
        {
            var oWorkflowstate = new ObjectVersionWorkflowState();
            var oLookup = new Lookup();
            var oObjID = new ObjID();

            if (iObjectTypeNoDocumento > 0) // Cualquier tipo de objeto que no sea documento
            {
                oObjID.SetIDs
                (
                    ObjType: iObjectTypeNoDocumento,
                    ID: oObjVer.ID
                );
            }
            else // Objeto tipo documento
            {
                oObjID.SetIDs
                (
                    ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                    ID: oObjVer.ID
                );
            }            

            var checkedOutObjectVersion = PermanentVault.ObjectOperations.GetLatestObjVerEx(oObjID, true);
            checkedOutObjectVersion = PermanentVault.ObjectOperations.CheckOut(oObjID).ObjVer;

            var oPropertyValue = new PropertyValue
            {
                PropertyDef = iPropertyDefEstatus
            };

            oLookup.Item = iEstatusIdValue;

            oPropertyValue.TypedValue.SetValueToLookup(oLookup);

            PermanentVault.ObjectPropertyOperations.SetProperty
            (
                ObjVer: checkedOutObjectVersion,
                PropertyValue: oPropertyValue
            );

            if (sTipoObjeto == "Documento")
            {
                if (iEstatusIdValue == 1) // Si el documento esta vigente, validar el estado de workflow para moverlo
                {
                    var documentoProperties = oObjVerEx.Properties;
                    var pdWorkflow = oObjVerEx.Vault.PropertyDefOperations.GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefWorkflow);
                    var pdState = oObjVerEx.Vault.PropertyDefOperations.GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefState);
                    var iWorkflowDocumento = documentoProperties.SearchForPropertyEx(pdWorkflow.ID, true).TypedValue.GetLookupID();
                    var iStateDocumento = documentoProperties.SearchForPropertyEx(pdState.ID, true).TypedValue.GetLookupID();

                    if (iWorkflowDocumento == iWorkflowValidacionesChecklist && iStateDocumento == iStateDocumentoProcesado)
                    {
                        // Actualizar el estado del workflow "Validaciones REPSE" dentro del documento
                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iWorkflow);
                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iState);
                        PermanentVault.ObjectPropertyOperations.SetWorkflowStateEx(checkedOutObjectVersion, oWorkflowstate);
                    }
                }
                else // Si el documento esta Vencido, siempre moverlo al estado Documento Vencido del Workflow Validaciones Checklist
                {
                    // Actualizar el estado del workflow "Validaciones REPSE" dentro del documento
                    oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iWorkflow);
                    oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iState);
                    PermanentVault.ObjectPropertyOperations.SetWorkflowStateEx(checkedOutObjectVersion, oWorkflowstate);
                }                                
            }            

            PermanentVault.ObjectOperations.CheckIn(checkedOutObjectVersion);
        }

        private void RelacionaDocumentosVigentes(
            int iDocumentosRelacionadosEnObjeto,
            ObjVerEx documentoProveedor,
            List<ObjVer> oListaDocumentosVigentes)
        {
            var oPropertyValues = documentoProveedor.Properties;
            var oPropertyValue = new PropertyValue();
            var oLookups = new Lookups();
            var oLookup = new Lookup();

            if (oPropertyValues.IndexOf(iDocumentosRelacionadosEnObjeto) != -1)
            {
                if (oListaDocumentosVigentes.Count > 0)
                {
                    foreach (ObjVer documento in oListaDocumentosVigentes)
                    {
                        oLookup.Item = documento.ID;
                        oLookups.Add(-1, oLookup);
                    }

                    var oObjVer = PermanentVault.ObjectOperations.GetLatestObjVerEx(documentoProveedor.ObjID, true);
                    oPropertyValue.PropertyDef = iDocumentosRelacionadosEnObjeto;
                    oPropertyValue.TypedValue.SetValueToMultiSelectLookup(oLookups);
                    oObjVer = PermanentVault.ObjectOperations.CheckOut(documentoProveedor.ObjID).ObjVer;
                    PermanentVault.ObjectPropertyOperations.SetProperty(oObjVer, oPropertyValue);
                    PermanentVault.ObjectOperations.CheckIn(oObjVer);
                }
            }
        }

        private DateTime ObtenerRangoDePeriodoDelDocumento(DateTime dtFechaInicio, string sVigencia, int iTipoRango)
        {
            DateTime dtFechaObtenida = new DateTime();
            int iAddValue = 0;

            switch (sVigencia)
            {
                case "Semanal":

                    if (iTipoRango == 1) iAddValue = 7;
                    else iAddValue = -7;
                    dtFechaObtenida = dtFechaInicio.AddDays(iAddValue);
                    break;

                case "Quincenal":

                    if (iTipoRango == 1) iAddValue = 15;
                    else iAddValue = -15;
                    dtFechaObtenida = dtFechaInicio.AddDays(iAddValue);
                    break;

                case "Mensual":

                    if (iTipoRango == 1) iAddValue = 1;
                    else iAddValue = -1;
                    dtFechaObtenida = dtFechaInicio.AddMonths(iAddValue);
                    break;

                case "Bimestral":

                    if (iTipoRango == 1) iAddValue = 2;
                    else iAddValue = -2;
                    dtFechaObtenida = dtFechaInicio.AddMonths(iAddValue);
                    break;

                case "Trimestral":

                    if (iTipoRango == 1) iAddValue = 3;
                    else iAddValue = -3;
                    dtFechaObtenida = dtFechaInicio.AddMonths(iAddValue);
                    break;

                case "Cuatrimestral":

                    if (iTipoRango == 1) iAddValue = 4;
                    else iAddValue = -4;
                    dtFechaObtenida = dtFechaInicio.AddMonths(iAddValue);
                    break;

                case "Anual":

                    if (iTipoRango == 1) iAddValue = 1;
                    else iAddValue = -1;
                    dtFechaObtenida = dtFechaInicio.AddYears(iAddValue);
                    break;
            }

            return dtFechaObtenida;
        }

        private string ObtenerPeriodoDeDocumentoFaltante(string sVigencia, DateTime dtFechaInicio, DateTime dtFechaFin)
        {
            string sPeriodoResultante = "";
            string sFechaInicio = "";

            switch (sVigencia)
            {
                case "No Aplica":
                    sPeriodoResultante = "Faltante";
                    break;

                case "Semanal":
                    sPeriodoResultante = dtFechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    break;

                case "Catorcenal":
                    sPeriodoResultante = dtFechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    break;

                case "Quincenal":
                    sPeriodoResultante = ObtenerFechaDePeriodoQuincenal(dtFechaInicio, dtFechaFin);
                    break;

                case "Mensual":
                    sPeriodoResultante = dtFechaInicio
                        .ToString("MMMM", new CultureInfo("es-ES")) + " " + dtFechaInicio
                        .ToString("yyyy", new CultureInfo("es-ES"));
                    break;

                case "Bimestral":
                    sFechaInicio = dtFechaInicio
                        .ToString("MMMM", new CultureInfo("es-ES")) + " " + dtFechaInicio
                        .ToString("yyyy", new CultureInfo("es-ES"));
                    sPeriodoResultante = sFechaInicio;
                    break;

                case "Trimestral":
                    sFechaInicio = dtFechaInicio
                        .ToString("MMMM", new CultureInfo("es-ES")) + " " + dtFechaInicio
                        .ToString("yyyy", new CultureInfo("es-ES"));
                    sPeriodoResultante = sFechaInicio;
                    break;

                case "Cuatrimestral":
                    sFechaInicio = dtFechaInicio
                        .ToString("MMMM", new CultureInfo("es-ES")) + " " + dtFechaInicio
                        .ToString("yyyy", new CultureInfo("es-ES"));
                    sPeriodoResultante = sFechaInicio;
                    break;

                case "Anual":
                    sPeriodoResultante = dtFechaInicio.ToString("yyyy", new CultureInfo("es-ES"));
                    break;
            }

            return sPeriodoResultante;
        }

        private string ObtenerFechaDePeriodoQuincenal(DateTime dtFechaInicio, DateTime dtFechaFin)
        {
            string sFechaResultante = "";
            List<DateTime> fechas = new List<DateTime>();

            // Extraer mes de fecha inicio y fin
            string sMesFechaInicio = dtFechaInicio.ToString("MMMM", new CultureInfo("es-ES"));
            string sMesFechaFin = dtFechaFin.ToString("MMMM", new CultureInfo("es-ES"));

            if (sMesFechaInicio == sMesFechaFin)
            {
                // Obtener 2 quincenas: 15 y ultimo dia del mes
                DateTime dtQuincena1 = new DateTime(dtFechaInicio.Year,dtFechaInicio.Month, 15);

                DateTime dtDia1DelMes = new DateTime(dtFechaInicio.Year, dtFechaInicio.Month, 1);
                DateTime dtQuincena2 = dtDia1DelMes.AddMonths(1).AddDays(-1);

                fechas.Add(dtQuincena1);
                fechas.Add(dtQuincena2);

                foreach (var fecha in fechas)
                {
                    var sFecha = fecha.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                    if (ComparaFechaBaseContraUnaFechaInicioYFechaFin(sFecha, dtFechaInicio, dtFechaFin) == true)
                    {
                        sFechaResultante = sFecha;
                    }
                }
            }
            else
            {
                // Si el Mes de fecha inicio y fin son diferentes, obtener 4 quincenas
                // 15 y ultimo dia de cada mes
                DateTime dtQuincena1 = new DateTime(dtFechaInicio.Year, dtFechaInicio.Month, 15);

                DateTime dtDia1DelMesQ2 = new DateTime(dtFechaInicio.Year, dtFechaInicio.Month, 1);
                DateTime dtQuincena2 = dtDia1DelMesQ2.AddMonths(1).AddDays(-1);

                DateTime dtQuincena3 = new DateTime(dtFechaFin.Year, dtFechaFin.Month, 15);

                DateTime dtDia1DelMesQ4 = new DateTime(dtFechaFin.Year, dtFechaFin.Month, 1);
                DateTime dtQuincena4 = dtDia1DelMesQ4.AddMonths(1).AddDays(-1);

                fechas.Add(dtQuincena1);
                fechas.Add(dtQuincena2);
                fechas.Add(dtQuincena3);
                fechas.Add(dtQuincena4);

                foreach (var fecha in fechas)
                {
                    var sFecha = fecha.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                    if (ComparaFechaBaseContraUnaFechaInicioYFechaFin(sFecha, dtFechaInicio, dtFechaFin) == true)
                    {
                        sFechaResultante = sFecha;
                    }
                }
            }

            return sFechaResultante;
        }

        private string ObtenerPeriodoDeVencimientoDelDocumento(DateTime dtFechaBaseOReferencia, string sVigenciaDocumento)
        {
            string sMesDeVencimiento = "";
            DateTime dtFechaFinDeVigencia = new DateTime();

            switch (sVigenciaDocumento)
            {
                case "Semanal":

                    // Sumar 7 dias a la fecha emision del documento y obtener el mes de la fecha resultante
                    dtFechaFinDeVigencia = dtFechaBaseOReferencia.AddDays(7);
                    sMesDeVencimiento = dtFechaFinDeVigencia.ToString("MMMM");

                    break;

                case "Mensual":

                    // Suma un mes para obtener fecha de fin de vigencia mensual del documento
                    dtFechaFinDeVigencia = dtFechaBaseOReferencia.AddMonths(1);
                    sMesDeVencimiento = dtFechaFinDeVigencia.ToString("MMMM");

                    break;

                case "Bimestral":

                    // Se obtiene el nombre del mes en que caduca el documento con vigencia bimestral
                    dtFechaFinDeVigencia = dtFechaBaseOReferencia.AddMonths(2);
                    sMesDeVencimiento = dtFechaFinDeVigencia.ToString("MMMM");

                    break;

                case "Trimestral":

                    // Se obtiene el nombre del mes en que caduca el documento con vigencia trimestral
                    dtFechaFinDeVigencia = dtFechaBaseOReferencia.AddMonths(3);
                    sMesDeVencimiento = dtFechaFinDeVigencia.ToString("MMMM");

                    break;

                case "Anual":

                    // Suma un año para obtener la fecha de vigencia anual del documento
                    dtFechaFinDeVigencia = dtFechaBaseOReferencia.AddYears(1);
                    sMesDeVencimiento = dtFechaFinDeVigencia.ToString("MMMM");

                    break;
            }

            return sMesDeVencimiento;
        }
    }
}
