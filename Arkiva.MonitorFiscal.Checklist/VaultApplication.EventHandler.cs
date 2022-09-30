using MFiles.VAF.Core;
using MFilesAPI;
using System;
using System.Collections.Generic;
using MFiles.VAF.Common;
using System.Globalization;

namespace Arkiva.MonitorFiscal.Checklist
{
    public partial class VaultApplication
        : ConfigurableVaultApplicationBase<Configuration>
    {
        #region Background Operation Task

        public void StartBackgroundOperationTask_2()
        {
            try
            {
                this.BackgroundOperations.StartRecurringBackgroundOperation(
                    "Recurring Background Operation",
                    TimeSpan.FromMinutes(4),
                () =>
                {
                    // Inicia el proceso de validacion de los documentos en el objeto
                    foreach (var grupo in Configuration.Grupos)
                    {
                        var pd_TipoDeValidacionLeyDeOutsourcing = PermanentVault
                            .PropertyDefOperations
                            .GetPropertyDefIDByAlias("PD.TipoDeValidacionLeyDeOutsourcing");

                        // busqueda de ojeto principal
                        var searchBuilderObject = new MFSearchBuilder(PermanentVault);
                        searchBuilderObject.Deleted(false);
                        searchBuilderObject.ObjType(grupo.ObjectType);

                        foreach (var objeto in searchBuilderObject.FindEx())
                        {
                            List<ObjVer> ListaDocumentosMasRecientesEnObjetoPadre = new List<ObjVer>();
                            //bool bExisteDocumentoRelacionadoNoActualizado = false;
                            bool bRelacionaDocumentosEnObjetoPadre = false;
                            var oPropertyValues = new PropertyValues();

                            oPropertyValues = PermanentVault
                                .ObjectPropertyOperations
                                .GetProperties(objeto.ObjVer);                            

                            if (oPropertyValues.IndexOf(pd_TipoDeValidacionLeyDeOutsourcing) != -1 && 
                                    !oPropertyValues.SearchForPropertyEx(pd_TipoDeValidacionLeyDeOutsourcing, true).TypedValue.IsNULL())
                            {
                                var tipoValidacionLeyOutsourcing = oPropertyValues
                                        .SearchForPropertyEx(pd_TipoDeValidacionLeyDeOutsourcing, true)
                                        .TypedValue
                                        .GetValueAsLocalizedText();

                                if (tipoValidacionLeyOutsourcing == "Por Proveedor")
                                {
                                    // Valida que exista la propiedad en la metadata del objeto
                                    if (oPropertyValues.IndexOf(grupo.MasRecientesDocumentosRelacionados) != -1)
                                    {
                                        var searchBuilderTiposDocumento = new MFSearchBuilder(PermanentVault);
                                        searchBuilderTiposDocumento.Deleted(false);
                                        searchBuilderTiposDocumento.ObjType(grupo.Checklist);

                                        // Recorrido por tipos de documento
                                        foreach (var tipoDocumento in searchBuilderTiposDocumento.FindEx())
                                        {
                                            var sComparaFecha1 = "";
                                            var sComparaFecha2 = "";
                                            var dtFecha1 = new DateTime();
                                            var dtFecha2 = new DateTime();
                                            var dtFechaFinal = new DateTime();
                                            var objVerDocumento1 = new ObjVer();
                                            var objVerDocumento2 = new ObjVer();
                                            var objVerDocumentoFinal = new ObjVer();
                                            bool bInicializarMetodoValidaVigenciaDocumentoTD = false;

                                            oPropertyValues = PermanentVault
                                                .ObjectPropertyOperations
                                                .GetProperties(tipoDocumento.ObjVer);

                                            if (oPropertyValues.IndexOf(grupo.CategoriaTipoDocumento) != -1 &&
                                                !oPropertyValues.SearchForPropertyEx(grupo.CategoriaTipoDocumento, true).TypedValue.IsNULL())
                                            {
                                                var categoriaTipoDocumento = oPropertyValues
                                                    .SearchForPropertyEx(grupo.CategoriaTipoDocumento, true)
                                                    .TypedValue
                                                    .GetValueAsLocalizedText();

                                                if (categoriaTipoDocumento == "Ley de Outsourcing")
                                                {
                                                    if (oPropertyValues.IndexOf(grupo.Vigencia) != -1 &&
                                                        !oPropertyValues.SearchForPropertyEx(grupo.Vigencia, true).TypedValue.IsNULL())
                                                    {
                                                        var vigenciaDocumentoTD = oPropertyValues
                                                            .SearchForPropertyEx(grupo.Vigencia.ID, true)
                                                            .TypedValue
                                                            .GetValueAsLocalizedText();

                                                        var searchBuilderDocumentosTD = new MFSearchBuilder(PermanentVault);
                                                        searchBuilderDocumentosTD.Deleted(false);
                                                        searchBuilderDocumentosTD.Property
                                                        (
                                                            grupo.PropertyDef,
                                                            MFDataType.MFDatatypeMultiSelectLookup,
                                                            objeto.ObjVer.ID
                                                        );
                                                        searchBuilderDocumentosTD.Property
                                                        (
                                                            grupo.PropertyDefChecklist,
                                                            MFDataType.MFDatatypeLookup,
                                                            tipoDocumento.ObjVer.ID
                                                        );

                                                        if (searchBuilderDocumentosTD.FindEx().Count > 0)
                                                        {
                                                            foreach (var documentoTD in searchBuilderDocumentosTD.FindEx())
                                                            {
                                                                oPropertyValues = PermanentVault
                                                                    .ObjectPropertyOperations
                                                                    .GetProperties(documentoTD.ObjVer);

                                                                if (oPropertyValues.IndexOf(grupo.FechaDeDocumento) != -1 &&
                                                                    !oPropertyValues.SearchForPropertyEx(grupo.FechaDeDocumento, true).TypedValue.IsNULL())
                                                                {
                                                                    var fechaEmisionDocumentoTD = oPropertyValues
                                                                        .SearchForProperty(grupo.FechaDeDocumento.ID)
                                                                        .TypedValue
                                                                        .Value;

                                                                    // Extraer fecha de emision de cada documento y compararlos
                                                                    // para extraer el mas reciente (tambien validar que este vigente)                           
                                                                    if (sComparaFecha1 == "")
                                                                    {
                                                                        dtFecha1 = Convert.ToDateTime(fechaEmisionDocumentoTD);
                                                                        sComparaFecha1 = dtFecha1.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                        objVerDocumento1 = documentoTD.ObjVer;
                                                                    }
                                                                    else // (sComparaFecha2 == "")
                                                                    {
                                                                        dtFecha2 = Convert.ToDateTime(fechaEmisionDocumentoTD);
                                                                        sComparaFecha2 = dtFecha2.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                        objVerDocumento2 = documentoTD.ObjVer;
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
                                                                            //objVerDocumento1 = null;
                                                                            objVerDocumentoFinal = objVerDocumento2;
                                                                            dtFechaFinal = dtFecha2;
                                                                        }
                                                                        else //if (iComparaFechasDeDocumentoChecklist > 0)
                                                                        {
                                                                            sComparaFecha2 = "";
                                                                            //objVerDocumento2 = null;
                                                                            objVerDocumentoFinal = objVerDocumento1;
                                                                            dtFechaFinal = dtFecha1;
                                                                        }

                                                                        bInicializarMetodoValidaVigenciaDocumentoTD = true;
                                                                    }

                                                                    // Si solo hay un documento, se establece el ID de objeto, la fecha de documento y la vigencia
                                                                    // directamente en el metodo "ValidaVigenciaDeDocumentosChecklistPorTipoOClaseDeDocumento"...
                                                                    if (searchBuilderDocumentosTD.FindEx().Count == 1)
                                                                    {
                                                                        objVerDocumentoFinal = objVerDocumento1;
                                                                        dtFechaFinal = dtFecha1;
                                                                        bInicializarMetodoValidaVigenciaDocumentoTD = true;
                                                                    }
                                                                }
                                                            }

                                                            if (bInicializarMetodoValidaVigenciaDocumentoTD)
                                                            {
                                                                if (!(vigenciaDocumentoTD == "No Aplica")) // Valida vigencia de documento
                                                                {
                                                                    // Validar vigencia del documento checklist mas reciente
                                                                    if (ValidaVigenciaDeDocumento(vigenciaDocumentoTD, dtFechaFinal) == true)
                                                                    {
                                                                        // Se obtienen solo los documentos mas recientes y que estan vigentes 
                                                                        ListaDocumentosMasRecientesEnObjetoPadre.Add(objVerDocumentoFinal);

                                                                        // Activa la relacion de objetos
                                                                        bRelacionaDocumentosEnObjetoPadre = true;
                                                                    }
                                                                }
                                                                else // Si el documento "No Aplica" validacion de vigencia
                                                                {
                                                                    ListaDocumentosMasRecientesEnObjetoPadre.Add(objVerDocumentoFinal);
                                                                    bRelacionaDocumentosEnObjetoPadre = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        // Recorrido por clases de documento
                                        foreach (var claseDocumento in grupo.DocumentosProveedor)
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

                                            var searchBuilderDocumentosCD = new MFSearchBuilder(PermanentVault);
                                            searchBuilderDocumentosCD.Deleted(false);
                                            searchBuilderDocumentosCD.Property
                                            (
                                                (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                                                MFDataType.MFDatatypeLookup,
                                                claseDocumento.DocumentoProveedor.ID
                                            );
                                            searchBuilderDocumentosCD.Property
                                            (
                                                grupo.PropertyDef,
                                                MFDataType.MFDatatypeMultiSelectLookup,
                                                objeto.ObjVer.ID
                                            );

                                            if (searchBuilderDocumentosCD.FindEx().Count > 0)
                                            {
                                                foreach (var documentoCD in searchBuilderDocumentosCD.FindEx())
                                                {
                                                    oPropertyValues = PermanentVault
                                                        .ObjectPropertyOperations
                                                        .GetProperties(documentoCD.ObjVer);

                                                    if (oPropertyValues.IndexOf(grupo.Vigencia) != -1 &&
                                                        oPropertyValues.IndexOf(grupo.FechaDeDocumento) != -1)
                                                    {
                                                        if (!oPropertyValues.SearchForPropertyEx(grupo.Vigencia, true).TypedValue.IsNULL() &&
                                                            !oPropertyValues.SearchForPropertyEx(grupo.FechaDeDocumento, true).TypedValue.IsNULL())
                                                        {              
                                                            // Validar estatus, solo comparar documentos vigentes

                                                            var fechaEmisionDocumentoCD = oPropertyValues
                                                                .SearchForProperty(grupo.FechaDeDocumento.ID)
                                                                .TypedValue
                                                                .Value;

                                                            // Extraer fecha de emision de cada documento y compararlos
                                                            // para extraer el mas reciente (tambien validar que este vigente)                           
                                                            if (sComparaFecha1 == "")
                                                            {
                                                                dtFecha1 = Convert.ToDateTime(fechaEmisionDocumentoCD);
                                                                sComparaFecha1 = dtFecha1.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                objVerDocumento1 = documentoCD.ObjVer;
                                                            }
                                                            else // (sComparaFecha2 == "")
                                                            {
                                                                dtFecha2 = Convert.ToDateTime(fechaEmisionDocumentoCD);
                                                                sComparaFecha2 = dtFecha2.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                                                                objVerDocumento2 = documentoCD.ObjVer;
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
                                                                    //objVerDocumento1 = null;
                                                                    objVerDocumentoFinal = objVerDocumento2;
                                                                    dtFechaFinal = dtFecha2;
                                                                }
                                                                else //if (iComparaFechasDeDocumentoChecklist > 0)
                                                                {
                                                                    sComparaFecha2 = "";
                                                                    //objVerDocumento2 = null;
                                                                    objVerDocumentoFinal = objVerDocumento1;
                                                                    dtFechaFinal = dtFecha1;
                                                                }

                                                                bInicializarMetodoValidaVigenciaDocumentoCD = true;
                                                            }

                                                            // Si solo hay un documento, se establece el ID de objeto, la fecha de documento y la vigencia
                                                            // directamente en el metodo "ValidaVigenciaDeDocumentosChecklistPorTipoOClaseDeDocumento"...
                                                            if (searchBuilderDocumentosCD.FindEx().Count == 1)
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
                                                        // Validar vigencia del documento checklist mas reciente
                                                        if (ValidaVigenciaDeDocumento(vigenciaDocumentoCD, dtFechaFinal) == true)
                                                        {
                                                            // Se obtienen solo los documentos mas recientes y que estan vigentes 
                                                            ListaDocumentosMasRecientesEnObjetoPadre.Add(objVerDocumentoFinal);

                                                            // Activa la relacion de objetos
                                                            bRelacionaDocumentosEnObjetoPadre = true;
                                                        }
                                                    }
                                                    else // Si el documento "No Aplica" validacion de vigencia
                                                    {
                                                        ListaDocumentosMasRecientesEnObjetoPadre.Add(objVerDocumentoFinal);
                                                        bRelacionaDocumentosEnObjetoPadre = true;
                                                    }
                                                }
                                            }
                                        }

                                        if (bRelacionaDocumentosEnObjetoPadre)
                                        {
                                            // Relaciona los documentos checklist mas recientes (y vigentes) en el objeto padre
                                            RelacionaDocumentosChecklistMasRecientesEnObjetoPadre(
                                                grupo.MasRecientesDocumentosRelacionados.ID,
                                                objeto,
                                                ListaDocumentosMasRecientesEnObjetoPadre);
                                        }
                                    }
                                }
                            }                            
                        }
                    }
                });

                SysUtils.ReportInfoToEventLog("Fin del proceso 'EventHandler'.");
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Ocurrio un error. " + ex.Message);
            }
        }

        #endregion

        #region Event Handler Methods

        //[EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges)]
        //protected void RelacionaDocumentosChecklistEnObjetoPadre(EventHandlerEnvironment env)
        //{
        //    // Inicia el proceso de validacion de los documentos en el objeto
        //    foreach (var grupo in Configuration.Grupos)
        //    {
        //        // busqueda de ojeto principal
        //        var searchBuilderObject = new MFSearchBuilder(PermanentVault);
        //        searchBuilderObject.Deleted(false);
        //        searchBuilderObject.ObjType(grupo.ObjectType);

        //        foreach (var objeto in searchBuilderObject.FindEx())
        //        {
        //            List<ObjVer> ListaDocumentosMasRecientesEnObjetoPadre = new List<ObjVer>();
        //            bool bExisteDocumentoRelacionadoNoActualizado = false;
        //            bool bRelacionaDocumentosEnObjetoPadre = false;
        //            var oPropertyValues = new PropertyValues();

        //            oPropertyValues = PermanentVault
        //                .ObjectPropertyOperations
        //                .GetProperties(objeto.ObjVer);

        //            // Valida que exista la propiedad en la metadata del objeto
        //            if (oPropertyValues.IndexOf(grupo.DocumentosRelacionadosEnObjeto) != -1)
        //            {
        //                // Valida si la propiedad contiene documentos relacionados
        //                if (!oPropertyValues.SearchForPropertyEx(grupo.DocumentosRelacionadosEnObjeto, true).TypedValue.IsNULL())
        //                {
        //                    var documentosRelacionados = oPropertyValues
        //                        .SearchForPropertyEx(grupo.DocumentosRelacionadosEnObjeto, true)
        //                        .TypedValue
        //                        .GetValueAsLookups().ToObjVerExs(PermanentVault);

        //                    foreach (var documentoRelacionado in documentosRelacionados)
        //                    {
        //                        // Buscar si los documentos adjuntos son los mas recientes
        //                        oPropertyValues = PermanentVault
        //                            .ObjectPropertyOperations
        //                            .GetProperties(documentoRelacionado.ObjVer);

        //                        if (oPropertyValues.IndexOf(grupo.FechaDeDocumento) != -1)
        //                        {
        //                            if (!oPropertyValues.SearchForPropertyEx(grupo.FechaDeDocumento, true).TypedValue.IsNULL())
        //                            {
        //                                var fechaEmisionDocumentoRelacionado = oPropertyValues
        //                                    .SearchForPropertyEx(grupo.FechaDeDocumento, true)
        //                                    .TypedValue
        //                                    .Value;

        //                                DateTime dtFechaEmisionDocumentoRelacionado = Convert.ToDateTime(fechaEmisionDocumentoRelacionado);

        //                                // Separar por clase y tipo de documento la validacion siguiente
        //                                // primero validar si el documento tiene la propiedad tipo de documento
        //                                // en caso de que no lo tenga es que es una clase ley outsourcing

        //                                // Si es busqueda por tipo de documento
        //                                if (oPropertyValues.IndexOf(grupo.PropertyDefChecklist) != -1)
        //                                {
        //                                    if (!oPropertyValues.SearchForPropertyEx(grupo.PropertyDefChecklist, true).TypedValue.IsNULL())
        //                                    {
        //                                        var oTipoDocumento = oPropertyValues
        //                                        .SearchForPropertyEx(grupo.PropertyDefChecklist, true)
        //                                        .TypedValue
        //                                        .GetValueAsLookup().ToObjVerEx(PermanentVault);

        //                                        var searchBuilderDocumentosTDExistentes = new MFSearchBuilder(PermanentVault);
        //                                        searchBuilderDocumentosTDExistentes.Deleted(false);
        //                                        searchBuilderDocumentosTDExistentes.Property
        //                                        (
        //                                            grupo.PropertyDef,
        //                                            MFDataType.MFDatatypeMultiSelectLookup,
        //                                            objeto.ObjVer.ID
        //                                        );
        //                                        searchBuilderDocumentosTDExistentes.Property
        //                                        (
        //                                            grupo.PropertyDefChecklist,
        //                                            MFDataType.MFDatatypeLookup,
        //                                            oTipoDocumento.ObjVer.ID
        //                                        );

        //                                        foreach (var documentoTDValidado in searchBuilderDocumentosTDExistentes.FindEx())
        //                                        {
        //                                            oPropertyValues = PermanentVault
        //                                                .ObjectPropertyOperations
        //                                                .GetProperties(documentoTDValidado.ObjVer);

        //                                            if (oPropertyValues.IndexOf(grupo.FechaDeDocumento) != -1)
        //                                            {
        //                                                if (!oPropertyValues.SearchForPropertyEx(grupo.FechaDeDocumento, true).TypedValue.IsNULL())
        //                                                {
        //                                                    var fechaEmisionDocumentoTDValidado = oPropertyValues
        //                                                        .SearchForPropertyEx(grupo.FechaDeDocumento, true)
        //                                                        .TypedValue
        //                                                        .Value;

        //                                                    DateTime dtFechaEmisionDocumentoTDValidado = Convert.ToDateTime(fechaEmisionDocumentoTDValidado);

        //                                                    if (documentoRelacionado.ObjVer.ID != documentoTDValidado.ObjVer.ID)
        //                                                    {
        //                                                        // Comparar fecha del documento relacionado con los otros de su mismo tipo
        //                                                        if (ValidaSiActualDocumentoRelacionadoEsElMasReciente(
        //                                                                dtFechaEmisionDocumentoRelacionado,
        //                                                                dtFechaEmisionDocumentoTDValidado) == false)
        //                                                        {
        //                                                            // Si se encuentra al menos un documento no actualizado se activa la bandera
        //                                                            bExisteDocumentoRelacionadoNoActualizado = true;
        //                                                        }
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                                else // Si es busqueda por clase de documento
        //                                {                                            
        //                                    //var oClaseDocumento = oPropertyValues
        //                                    //    .SearchForPropertyEx((int)MFBuiltInPropertyDef
        //                                    //    .MFBuiltInPropertyDefClass, true)
        //                                    //    .TypedValue
        //                                    //    .GetValueAsLookup().ToObjVerEx(PermanentVault);

        //                                    if (documentosRelacionados.Count == 10)
        //                                    {
        //                                        var iClaseDocumento = oPropertyValues
        //                                        .SearchForPropertyEx((int)MFBuiltInPropertyDef
        //                                        .MFBuiltInPropertyDefClass, true)
        //                                        .TypedValue
        //                                        .GetLookupID();

        //                                        if (iClaseDocumento > 0) //foreach (var claseDocumento in grupo.ClasesChecklist)
        //                                        {
        //                                            var searchBuilderDocumentosCDExistentes = new MFSearchBuilder(PermanentVault);
        //                                            searchBuilderDocumentosCDExistentes.Deleted(false);
        //                                            searchBuilderDocumentosCDExistentes.Property
        //                                            (
        //                                                grupo.PropertyDef,
        //                                                MFDataType.MFDatatypeMultiSelectLookup,
        //                                                objeto.ObjVer.ID
        //                                            );
        //                                            //searchBuilderDocumentosCDExistentes.Property
        //                                            //(
        //                                            //    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
        //                                            //    MFDataType.MFDatatypeLookup,
        //                                            //    claseDocumento.ChecklistClass.ID
        //                                            //);
        //                                            searchBuilderDocumentosCDExistentes.Class(iClaseDocumento);

        //                                            foreach (var documentoCDValidado in searchBuilderDocumentosCDExistentes.FindEx())
        //                                            {
        //                                                oPropertyValues = PermanentVault
        //                                                    .ObjectPropertyOperations
        //                                                    .GetProperties(documentoCDValidado.ObjVer);

        //                                                if (oPropertyValues.IndexOf(grupo.FechaDeDocumento) != -1)
        //                                                {
        //                                                    if (!oPropertyValues.SearchForPropertyEx(grupo.FechaDeDocumento, true).TypedValue.IsNULL())
        //                                                    {
        //                                                        var fechaEmisionDocumentoCDValidado = oPropertyValues
        //                                                            .SearchForPropertyEx(grupo.FechaDeDocumento, true)
        //                                                            .TypedValue
        //                                                            .Value;

        //                                                        DateTime dtFechaEmisionDocumentoCDValidado = Convert.ToDateTime(fechaEmisionDocumentoCDValidado);

        //                                                        if (documentoRelacionado.ObjVer.ID != documentoCDValidado.ObjVer.ID)
        //                                                        {
        //                                                            // Comparar fecha del documento relacionado con los otros de su misma clase
        //                                                            if (ValidaSiActualDocumentoRelacionadoEsElMasReciente(
        //                                                                    dtFechaEmisionDocumentoRelacionado,
        //                                                                    dtFechaEmisionDocumentoCDValidado) == false)
        //                                                            {
        //                                                                // Si se encuentra al menos un documento no actualizado se activa la bandera
        //                                                                bExisteDocumentoRelacionadoNoActualizado = true;
        //                                                            }
        //                                                        }
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        // Si existen menos de 10 documentos relacionados en el objeto principal se activa la bandera
        //                                        bExisteDocumentoRelacionadoNoActualizado = true;
        //                                    }                                            
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    // Si la propiedad se encuentra sin documentos relacionados, tambien se activa la bandera
        //                    bExisteDocumentoRelacionadoNoActualizado = true;
        //                }
        //            }

        //            if (bExisteDocumentoRelacionadoNoActualizado)
        //            {
        //                var searchBuilderTiposDocumento = new MFSearchBuilder(PermanentVault);
        //                searchBuilderTiposDocumento.Deleted(false);
        //                searchBuilderTiposDocumento.ObjType(grupo.Checklist);

        //                // Recorrido por tipos de documento
        //                foreach (var tipoDocumento in searchBuilderTiposDocumento.FindEx())
        //                {
        //                    var sComparaFecha1 = "";
        //                    var sComparaFecha2 = "";
        //                    var dtFecha1 = new DateTime();
        //                    var dtFecha2 = new DateTime();
        //                    var dtFechaFinal = new DateTime();
        //                    var objVerDocumento1 = new ObjVer();
        //                    var objVerDocumento2 = new ObjVer();
        //                    var objVerDocumentoFinal = new ObjVer();
        //                    bool bInicializarMetodoValidaVigenciaDocumentoTD = false;

        //                    oPropertyValues = PermanentVault
        //                        .ObjectPropertyOperations
        //                        .GetProperties(tipoDocumento.ObjVer);

        //                    if (oPropertyValues.IndexOf(grupo.CategoriaTipoDocumento) != -1)
        //                    {
        //                        if (!oPropertyValues.SearchForPropertyEx(grupo.CategoriaTipoDocumento, true).TypedValue.IsNULL())
        //                        {
        //                            var categoriaTipoDocumento = oPropertyValues
        //                                .SearchForPropertyEx(grupo.CategoriaTipoDocumento, true)
        //                                .TypedValue
        //                                .GetValueAsLocalizedText();

        //                            if (categoriaTipoDocumento == "Ley de Outsourcing")
        //                            {
        //                                if (oPropertyValues.IndexOf(grupo.Vigencia) != -1)
        //                                {
        //                                    if (!oPropertyValues.SearchForPropertyEx(grupo.Vigencia, true).TypedValue.IsNULL())
        //                                    {
        //                                        var vigenciaDocumentoTD = oPropertyValues
        //                                            .SearchForPropertyEx(grupo.Vigencia.ID, true)
        //                                            .TypedValue
        //                                            .GetValueAsLocalizedText();

        //                                        var searchBuilderDocumentosTD = new MFSearchBuilder(PermanentVault);
        //                                        searchBuilderDocumentosTD.Deleted(false);
        //                                        searchBuilderDocumentosTD.Property
        //                                        (
        //                                            grupo.PropertyDef,
        //                                            MFDataType.MFDatatypeMultiSelectLookup,
        //                                            objeto.ObjVer.ID
        //                                        );
        //                                        searchBuilderDocumentosTD.Property
        //                                        (
        //                                            grupo.PropertyDefChecklist,
        //                                            MFDataType.MFDatatypeLookup,
        //                                            tipoDocumento.ObjVer.ID
        //                                        );

        //                                        if (searchBuilderDocumentosTD.FindEx().Count > 0)
        //                                        {
        //                                            foreach (var documentoTD in searchBuilderDocumentosTD.FindEx())
        //                                            {
        //                                                oPropertyValues = PermanentVault
        //                                                    .ObjectPropertyOperations
        //                                                    .GetProperties(documentoTD.ObjVer);

        //                                                if (oPropertyValues.IndexOf(grupo.FechaDeDocumento) != -1)
        //                                                {
        //                                                    if (!oPropertyValues.SearchForPropertyEx(grupo.FechaDeDocumento, true).TypedValue.IsNULL())
        //                                                    {
        //                                                        var fechaEmisionDocumentoTD = oPropertyValues
        //                                                            .SearchForProperty(grupo.FechaDeDocumento.ID)
        //                                                            .TypedValue
        //                                                            .Value;

        //                                                        // Extraer fecha de emision de cada documento y compararlos
        //                                                        // para extraer el mas reciente (tambien validar que este vigente)                           
        //                                                        if (sComparaFecha1 == "")
        //                                                        {
        //                                                            dtFecha1 = Convert.ToDateTime(fechaEmisionDocumentoTD);
        //                                                            sComparaFecha1 = dtFecha1.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        //                                                            objVerDocumento1 = documentoTD.ObjVer;
        //                                                        }
        //                                                        else // (sComparaFecha2 == "")
        //                                                        {
        //                                                            dtFecha2 = Convert.ToDateTime(fechaEmisionDocumentoTD);
        //                                                            sComparaFecha2 = dtFecha2.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        //                                                            objVerDocumento2 = documentoTD.ObjVer;
        //                                                        }

        //                                                        if (sComparaFecha1 != "" && sComparaFecha2 != "")
        //                                                        {
        //                                                            int iComparaFechasDeDocumentoChecklist = DateTime.Compare
        //                                                            (
        //                                                                Convert.ToDateTime(sComparaFecha1),
        //                                                                Convert.ToDateTime(sComparaFecha2)
        //                                                            );

        //                                                            if (iComparaFechasDeDocumentoChecklist < 0)
        //                                                            {
        //                                                                sComparaFecha1 = "";
        //                                                                //objVerDocumento1 = null;
        //                                                                objVerDocumentoFinal = objVerDocumento2;
        //                                                                dtFechaFinal = dtFecha2;
        //                                                            }
        //                                                            else //if (iComparaFechasDeDocumentoChecklist > 0)
        //                                                            {
        //                                                                sComparaFecha2 = "";
        //                                                                //objVerDocumento2 = null;
        //                                                                objVerDocumentoFinal = objVerDocumento1;
        //                                                                dtFechaFinal = dtFecha1;
        //                                                            }

        //                                                            bInicializarMetodoValidaVigenciaDocumentoTD = true;
        //                                                        }

        //                                                        // Si solo hay un documento, se establece el ID de objeto, la fecha de documento y la vigencia
        //                                                        // directamente en el metodo "ValidaVigenciaDeDocumentosChecklistPorTipoOClaseDeDocumento"...
        //                                                        if (searchBuilderDocumentosTD.FindEx().Count == 1)
        //                                                        {
        //                                                            objVerDocumentoFinal = objVerDocumento1;
        //                                                            dtFechaFinal = dtFecha1;
        //                                                            bInicializarMetodoValidaVigenciaDocumentoTD = true;
        //                                                        }
        //                                                    }
        //                                                }
        //                                            }

        //                                            if (bInicializarMetodoValidaVigenciaDocumentoTD)
        //                                            {
        //                                                if (!(vigenciaDocumentoTD == "No Aplica")) // Valida vigencia de documento
        //                                                {
        //                                                    // Validar vigencia del documento checklist mas reciente
        //                                                    if (ValidaVigenciaDeDocumentosChecklistPorTipoOClaseDeDocumento(
        //                                                            grupo.ClaseDocumentoRelacionadoEnChecklist.ID,
        //                                                            grupo.FechaDeDocumento.ID,
        //                                                            objVerDocumentoFinal,
        //                                                            dtFechaFinal,
        //                                                            vigenciaDocumentoTD) == true)
        //                                                    {
        //                                                        // Se obtienen solo los documentos mas recientes y que estan vigentes 
        //                                                        ListaDocumentosMasRecientesEnObjetoPadre.Add(objVerDocumentoFinal);

        //                                                        // Activa la relacion de objetos
        //                                                        bRelacionaDocumentosEnObjetoPadre = true;
        //                                                    }
        //                                                }
        //                                                else // Si el documento "No Aplica" validacion de vigencia
        //                                                {
        //                                                    ListaDocumentosMasRecientesEnObjetoPadre.Add(objVerDocumentoFinal);
        //                                                    bRelacionaDocumentosEnObjetoPadre = true;
        //                                                }                                                        
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }
        //                }

        //                // Recorrido por clases de documento
        //                foreach (var claseDocumento in grupo.ClasesChecklist)
        //                {
        //                    var sComparaFecha1 = "";
        //                    var sComparaFecha2 = "";
        //                    var dtFecha1 = new DateTime();
        //                    var dtFecha2 = new DateTime();
        //                    var dtFechaFinal = new DateTime();
        //                    var objVerDocumento1 = new ObjVer();
        //                    var objVerDocumento2 = new ObjVer();
        //                    var objVerDocumentoFinal = new ObjVer();
        //                    var vigenciaDocumentoCD = "";
        //                    bool bInicializarMetodoValidaVigenciaDocumentoCD = false;

        //                    var searchBuilderDocumentosCD = new MFSearchBuilder(PermanentVault);
        //                    searchBuilderDocumentosCD.Deleted(false);
        //                    searchBuilderDocumentosCD.Property
        //                    (
        //                        (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
        //                        MFDataType.MFDatatypeLookup,
        //                        claseDocumento.ChecklistClass.ID
        //                    );
        //                    searchBuilderDocumentosCD.Property
        //                    (
        //                        grupo.PropertyDef,
        //                        MFDataType.MFDatatypeMultiSelectLookup,
        //                        objeto.ObjVer.ID
        //                    );

        //                    if (searchBuilderDocumentosCD.FindEx().Count > 0)
        //                    {
        //                        foreach (var documentoCD in searchBuilderDocumentosCD.FindEx())
        //                        {
        //                            oPropertyValues = PermanentVault
        //                                .ObjectPropertyOperations
        //                                .GetProperties(documentoCD.ObjVer);

        //                            if (oPropertyValues.IndexOf(grupo.Vigencia) != -1 &&
        //                                oPropertyValues.IndexOf(grupo.FechaDeDocumento) != -1)
        //                            {
        //                                if (!oPropertyValues.SearchForPropertyEx(grupo.Vigencia, true).TypedValue.IsNULL() &&
        //                                !oPropertyValues.SearchForPropertyEx(grupo.FechaDeDocumento, true).TypedValue.IsNULL())
        //                                {
        //                                    // Obtener Vigencia de la clase verificada
        //                                    vigenciaDocumentoCD = oPropertyValues
        //                                        .SearchForPropertyEx(grupo.Vigencia.ID, true)
        //                                        .TypedValue
        //                                        .GetValueAsLocalizedText();

        //                                    var fechaEmisionDocumentoCD = oPropertyValues
        //                                        .SearchForProperty(grupo.FechaDeDocumento.ID)
        //                                        .TypedValue
        //                                        .Value;

        //                                    // Extraer fecha de emision de cada documento y compararlos
        //                                    // para extraer el mas reciente (tambien validar que este vigente)                           
        //                                    if (sComparaFecha1 == "")
        //                                    {
        //                                        dtFecha1 = Convert.ToDateTime(fechaEmisionDocumentoCD);
        //                                        sComparaFecha1 = dtFecha1.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        //                                        objVerDocumento1 = documentoCD.ObjVer;
        //                                    }
        //                                    else // (sComparaFecha2 == "")
        //                                    {
        //                                        dtFecha2 = Convert.ToDateTime(fechaEmisionDocumentoCD);
        //                                        sComparaFecha2 = dtFecha2.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        //                                        objVerDocumento2 = documentoCD.ObjVer;
        //                                    }

        //                                    if (sComparaFecha1 != "" && sComparaFecha2 != "")
        //                                    {
        //                                        int iComparaFechasDeDocumentoChecklist = DateTime.Compare
        //                                        (
        //                                            Convert.ToDateTime(sComparaFecha1),
        //                                            Convert.ToDateTime(sComparaFecha2)
        //                                        );

        //                                        if (iComparaFechasDeDocumentoChecklist < 0)
        //                                        {
        //                                            sComparaFecha1 = "";
        //                                            //objVerDocumento1 = null;
        //                                            objVerDocumentoFinal = objVerDocumento2;
        //                                            dtFechaFinal = dtFecha2;
        //                                        }
        //                                        else //if (iComparaFechasDeDocumentoChecklist > 0)
        //                                        {
        //                                            sComparaFecha2 = "";
        //                                            //objVerDocumento2 = null;
        //                                            objVerDocumentoFinal = objVerDocumento1;
        //                                            dtFechaFinal = dtFecha1;
        //                                        }

        //                                        bInicializarMetodoValidaVigenciaDocumentoCD = true;
        //                                    }

        //                                    // Si solo hay un documento, se establece el ID de objeto, la fecha de documento y la vigencia
        //                                    // directamente en el metodo "ValidaVigenciaDeDocumentosChecklistPorTipoOClaseDeDocumento"...
        //                                    if (searchBuilderDocumentosCD.FindEx().Count == 1)
        //                                    {
        //                                        objVerDocumentoFinal = objVerDocumento1;
        //                                        dtFechaFinal = dtFecha1;
        //                                        bInicializarMetodoValidaVigenciaDocumentoCD = true;
        //                                    }
        //                                }
        //                            }
        //                        }

        //                        if (bInicializarMetodoValidaVigenciaDocumentoCD)
        //                        {
        //                            if (!(vigenciaDocumentoCD == "No Aplica")) // Valida vigencia de documento
        //                            {
        //                                // Validar vigencia del documento checklist mas reciente
        //                                if (ValidaVigenciaDeDocumentosChecklistPorTipoOClaseDeDocumento(
        //                                        grupo.ClaseDocumentoRelacionadoEnChecklist.ID,
        //                                        grupo.FechaDeDocumento.ID,
        //                                        objVerDocumentoFinal,
        //                                        dtFechaFinal,
        //                                        vigenciaDocumentoCD) == true)
        //                                {
        //                                    // Se obtienen solo los documentos mas recientes y que estan vigentes 
        //                                    ListaDocumentosMasRecientesEnObjetoPadre.Add(objVerDocumentoFinal);

        //                                    // Activa la relacion de objetos
        //                                    bRelacionaDocumentosEnObjetoPadre = true;
        //                                }
        //                            }
        //                            else // Si el documento "No Aplica" validacion de vigencia
        //                            {
        //                                ListaDocumentosMasRecientesEnObjetoPadre.Add(objVerDocumentoFinal);
        //                                bRelacionaDocumentosEnObjetoPadre = true;
        //                            }                                    
        //                        }
        //                    }
        //                }

        //                if (bRelacionaDocumentosEnObjetoPadre)
        //                {
        //                    // Relaciona los documentos checklist mas recientes (y vigentes) en el objeto padre
        //                    RelacionaDocumentosChecklistMasRecientesEnObjetoPadre(
        //                        grupo.DocumentosRelacionadosEnObjeto.ID,
        //                        objeto,
        //                        ListaDocumentosMasRecientesEnObjetoPadre);
        //                }
        //            }
        //        }
        //    }
        //}

        #endregion

        #region Metodos y Rutinas

        private bool ValidaSiActualDocumentoRelacionadoEsElMasReciente(
            DateTime dtFechaDocumentoRelacionado, 
            DateTime dtFechaDocumentoAComparar)
        {
            bool bResultado = false; // El resultado se inicializa en false

            string sFechaUno = dtFechaDocumentoRelacionado.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            string sFechaDos = dtFechaDocumentoAComparar.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            int iComparaFechaDeDocumentoRelacionado = DateTime.Compare
            (
                Convert.ToDateTime(sFechaUno), 
                Convert.ToDateTime(sFechaDos)
            );

            if (iComparaFechaDeDocumentoRelacionado >= 0)
            {
                bResultado = true;
            }

            return bResultado;
        }

        private bool ValidaVigenciaDeDocumento(string sVigenciaDocumento, DateTime dtFechaDocumento)
        {
            bool bDocumentoVigente = false;

            DateTime dtFechaActual = DateTime.Today;

            switch (sVigenciaDocumento)
            {
                case "Semanal":

                    if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist_EventHandler(
                        dtFechaDocumento,
                        dtFechaActual) <= 7)
                    {
                        bDocumentoVigente = true;
                    }

                    break;

                case "Mensual":

                    if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist_EventHandler(
                        dtFechaDocumento,
                        dtFechaActual) <= 30)
                    {
                        bDocumentoVigente = true;
                    }

                    break;

                case "Bimestral":

                    if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist_EventHandler(
                        dtFechaDocumento,
                        dtFechaActual) <= 60)
                    {
                        bDocumentoVigente = true;
                    }

                    break;

                case "Trimestral":

                    if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist_EventHandler(
                        dtFechaDocumento,
                        dtFechaActual) <= 90)
                    {
                        bDocumentoVigente = true;
                    }

                    break;

                case "Anual":

                    if (CalculaDiasEntreDocumentoPadreYDocumentosChecklist_EventHandler(
                        dtFechaDocumento,
                        dtFechaActual) <= 365)
                    {
                        bDocumentoVigente = true;
                    }

                    break;
            }         
            
            return bDocumentoVigente;
        }

        private void RelacionaDocumentosChecklistMasRecientesEnObjetoPadre(
            int iDocumentosRelacionadosEnObjeto,
            ObjVerEx objeto,
            List<ObjVer> listaDeDocumentosMasRecientesDelObjetoPadre)
        {
            var oPropertyValues = objeto.Properties;
            var oPropertyValue = new PropertyValue();
            var oLookups = new Lookups();
            var oLookup = new Lookup();

            if (oPropertyValues.IndexOf(iDocumentosRelacionadosEnObjeto) != -1)
            {
                if (listaDeDocumentosMasRecientesDelObjetoPadre.Count > 0)
                {
                    foreach (ObjVer documento in listaDeDocumentosMasRecientesDelObjetoPadre)
                    {
                        oLookup.Item = documento.ID;
                        oLookups.Add(-1, oLookup);
                    }

                    var oObjVer = PermanentVault.ObjectOperations.GetLatestObjVerEx(objeto.ObjID, true);
                    oPropertyValue.PropertyDef = iDocumentosRelacionadosEnObjeto;
                    oPropertyValue.TypedValue.SetValueToMultiSelectLookup(oLookups);
                    oObjVer = PermanentVault.ObjectOperations.CheckOut(objeto.ObjID).ObjVer;
                    PermanentVault.ObjectPropertyOperations.SetProperty(oObjVer, oPropertyValue);
                    PermanentVault.ObjectOperations.CheckIn(oObjVer);
                }
            }
        }

        private int CalculaDiasEntreDocumentoPadreYDocumentosChecklist_EventHandler(
            DateTime dtFechaEmisionDocumento,
            DateTime dtFechaEmisionChecklist)
        {
            TimeSpan diferenciaDeFechas = new TimeSpan();
            int diferenciaEnDias = 0;

            string sFecha1 = dtFechaEmisionDocumento.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            string sFecha2 = dtFechaEmisionChecklist.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            int iComparaFechasDeEmision = DateTime.Compare
            (
                Convert.ToDateTime(sFecha1),
                Convert.ToDateTime(sFecha2)
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

        #endregion
    }
}
