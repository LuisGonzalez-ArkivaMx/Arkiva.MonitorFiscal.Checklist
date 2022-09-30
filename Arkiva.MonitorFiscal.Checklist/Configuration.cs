using MFiles.VAF.Configuration;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Arkiva.MonitorFiscal.Checklist
{
    [DataContract]
    public class Configuration
    {
        // NOTE: The default value needs to be placed in both the JsonConfEditor
        // (or derived) attribute, and as a default value on the member.        

        [MFPropertyDef]
        [JsonConfEditor(
            Label = "Contacto",
            HelpText = "Propiedad de Contacto que se define en el objeto padre")]
        [DataMember]
        public MFIdentifier ContactoExterno { get; set; }

        [MFPropertyDef]
        [JsonConfEditor(
            Label = "E-mail",
            HelpText = "Propiedad de E-mail que se define en el objeto contacto externo/interno")]
        [DataMember]
        public MFIdentifier CorreoElectronico { get; set; }

        [DataMember]
        [JsonConfEditor(
            Label = "E-mail de remitente",
            HelpText = "Agregar el correo desde la cual se enviaran las notificaciones/e-mails")]
        [Security(ChangeBy = SecurityAttribute.UserLevel.VaultAdmin)]
        public string RemitenteEmail { get; set; }

        [DataMember]
        [JsonConfEditor(
            Label = "Tiempo de espera (en mins)",
            HelpText = "Definir el tiempo de espera (establecido en minutos) entre las ejecuciones de la aplicacion")]
        [Security(ChangeBy = SecurityAttribute.UserLevel.VaultAdmin)]
        [JsonConfIntegerEditor]
        public int IntervaloDeEjecucionEnMins { get; set; }

        [DataMember]
        public List<Grupo> Grupos { get; set; }
    }

    [DataContract]
    public class Grupo
    {
        // Nombre de grupo
        [DataMember]
        [TextEditor(IsRequired = true)]
        public string Name { get; set; }

        [MFObjType]
        [JsonConfEditor(Label = "Objeto Proveedor")]
        [DataMember]
        public MFIdentifier ObjectType { get; set; }

        [MFPropertyDef]
        [JsonConfEditor(
            Label = "Propiedad Proveedor SE Documental",
            HelpText = "Esta propiedad se visualiza en las clases de referencia y documentos que los relaciona al objeto origen")]
        [DataMember]
        public MFIdentifier PropertyDefProveedorSEDocumentos { get; set; }

        [DataMember]
        [MFPropertyDef(AllowEmpty = true)]
        [JsonConfEditor(IsRequired = false, DefaultValue = null, Label = "Propiedad Proveedor SE No Documental")]
        public MFIdentifier PropertyDefProveedorSENoDocumentos { get; set; }

        [DataMember]
        [MFObjType(AllowEmpty = true)]
        [JsonConfEditor(IsRequired = false, DefaultValue = null, Label = "Objeto de tipo documento")]
        public MFIdentifier Checklist { get; set; }

        [DataMember]
        [MFPropertyDef(AllowEmpty = true)]
        [JsonConfEditor(IsRequired = false, DefaultValue = null, Label = "Propiedad de tipo documento")]
        public MFIdentifier PropertyDefChecklist { get; set; }

        [DataMember]
        [MFPropertyDef(AllowEmpty = true)]
        [JsonConfEditor(IsRequired = false, DefaultValue = null, Label = "Categoria de tipo documento")]
        public MFIdentifier CategoriaTipoDocumento { get; set; }

        [MFPropertyDef]
        [JsonConfEditor(Label = "Fecha de documento")]
        [DataMember]
        public MFIdentifier FechaDeDocumento { get; set; }

        [MFPropertyDef]
        [JsonConfEditor(Label = "Fecha de inicio (Proveedor)")]
        [DataMember]
        public MFIdentifier FechaInicioProveedor { get; set; }

        [MFPropertyDef]
        [JsonConfEditor(Label = "Vigencia de documento")]
        [DataMember]
        public MFIdentifier Vigencia { get; set; }

        [MFPropertyDef]
        [JsonConfEditor(Label = "Propiedad para relacion de documentos mas recientes")]
        [DataMember]
        public MFIdentifier MasRecientesDocumentosRelacionados { get; set; }

        [MFPropertyDef]
        [JsonConfEditor(Label = "Propiedad para relacion de historial de documentos")]
        [DataMember]
        public MFIdentifier HistorialDocumentosRelacionados { get; set; }

        [MFPropertyDef]
        [JsonConfEditor(Label = "Propiedad que activa la ley")]
        [DataMember]
        public MFIdentifier CheckboxLeyOutsourcing { get; set; }

        [DataMember]
        [Security(ChangeBy = SecurityAttribute.UserLevel.VaultAdmin)]
        [JsonConfEditor(TypeEditor = "date", Label = "Fecha inicio de ley")]
        public DateTime? FechaInicio { get; set; }

        [DataMember]
        [Security(ChangeBy = SecurityAttribute.UserLevel.VaultAdmin)]
        [JsonConfEditor(TypeEditor = "date", Label = "Fecha fin de ley")]
        public DateTime? FechaFin { get; set; }

        [DataMember]
        [JsonConfEditor(Label = "Clases de referencia")]
        public List<Clases_Referencia> ClasesReferencia { get; set; }

        [DataMember]
        [JsonConfEditor(Label = "Documentos proveedor")]
        public List<Documentos_Proveedor> DocumentosProveedor { get; set; }

        [DataMember]
        [JsonConfEditor(Label = "Documentos empleado")]
        public List<Documentos_Empleado> DocumentosEmpleado { get; set; }

        [DataMember]
        [JsonConfEditor(Label = "Flujo de trabajo")]
        public ConfigurationWorkflow ConfigurationWorkflow { get; set; }

    }

    [DataContract]
    [JsonConfEditor(NameMember = "ClaseReferencia")]
    public class Clases_Referencia
    {
        [MFClass]
        [JsonConfEditor(Label = "Clase")]
        [DataMember]
        [TextEditor(IsRequired = true)]
        public MFIdentifier ClaseReferencia { get; set; }
    }

    [DataContract]
    [JsonConfEditor(NameMember = "DocumentoProveedor")]
    public class Documentos_Proveedor
    {
        [DataMember]
        [JsonConfEditor(
            Label = "Nombre",
            HelpText = "Nombre de Documento del Proveedor, requerido para el envio del correo o notificacion")]
        [TextEditor(IsRequired = true)]
        public string NombreClaseDocumento { get; set; }

        [MFClass]
        [JsonConfEditor(Label = "Clase")]
        [DataMember]
        [TextEditor(IsRequired = true)]
        public MFIdentifier DocumentoProveedor { get; set; }

        [DataMember]
        [JsonConfEditor(
            Label = "Vigencia",
            TypeEditor = "options",
            Options = "{selectOptions:[\"Anual\",\"Bimestral\",\"Mensual\",\"No Aplica\",\"Semanal\",\"Trimestral\",\"Cuatrimestral\"]}")]
        public string VigenciaDocumentoProveedor;

        [DataMember]
        [MFPropertyDef(AllowEmpty = true)]
        [JsonConfEditor(IsRequired = false, DefaultValue = null, Label = "Propiedad Documento Relacionado")]
        public MFIdentifier PropertyDefDocumentoRelacionado { get; set; }

        [DataMember]
        [JsonConfEditor(
            Label = "Tipo de documento",
            TypeEditor = "options",
            Options = "{selectOptions:[\"Documento checklist\",\"Comprobante de pago\"]}")]
        public string TipoDocumentoChecklist;
    }

    [DataContract]
    [JsonConfEditor(NameMember = "DocumentoEmpleado")]
    public class Documentos_Empleado
    {
        [DataMember]
        [JsonConfEditor(
            Label = "Nombre",
            HelpText = "Nombre de la clase documento del Empleado, requerido para el envio del correo o notificacion")]
        [TextEditor(IsRequired = true)]
        public string NombreClaseDocumento { get; set; }

        [MFClass]
        [JsonConfEditor(Label = "Clase")]
        [DataMember]
        [TextEditor(IsRequired = true)]
        public MFIdentifier DocumentoEmpleado { get; set; }

        [DataMember]
        [JsonConfEditor(
            Label = "Tipo de Validacion",
            TypeEditor = "options",
            Options = "{selectOptions:[\"Por empleado\",\"Por frecuencia de pago\"]}")]
        public string TipoValidacion;
    }

    [DataContract]
    public class ConfigurationWorkflow
    {
        [DataMember]
        [JsonConfEditor(Label = "Validaciones Checklist")]
        public WorkflowChecklist WorkflowChecklist { get; set; }

        [DataMember]
        [JsonConfEditor(Label = "Validaciones Documento Proveedor")]
        public WorkflowDocumentoProveedor WorkflowDocumentoProveedor { get; set; }

        [DataMember]
        [JsonConfEditor(Label = "Validaciones Documento Empleado")]
        public WorkflowDocumentoEmpleado WorkflowDocumentoEmpleado { get; set; }
    }

    [DataContract]
    public class WorkflowChecklist
    {
        [MFWorkflow]
        [JsonConfEditor(Label = "Flujo de trabajo")]
        [DataMember]
        public MFIdentifier WorkflowValidacionesChecklist { get; set; }

        [MFState]
        [DataMember]
        [JsonConfEditor(Label = "Documento Procesado")]
        public MFIdentifier EstadoDocumentoProcesado { get; set; }

        [MFState]
        [DataMember]
        [JsonConfEditor(Label = "Documento Vigente")]
        public MFIdentifier EstadoDocumentoVigente { get; set; }

        [MFState]
        [DataMember]
        [JsonConfEditor(Label = "Documento Vencido")]
        public MFIdentifier EstadoDocumentoVencido { get; set; }
    }

    [DataContract]
    public class WorkflowDocumentoProveedor
    {
        [MFWorkflow]
        [JsonConfEditor(Label = "Flujo de trabajo")]
        [DataMember]
        public MFIdentifier WorkflowValidacionesDocProveedor { get; set; }

        [MFState]
        [DataMember]
        [JsonConfEditor(Label = "Documento Vigente")]
        public MFIdentifier EstadoDocumentoVigenteProveedor { get; set; }
    }

    [DataContract]
    public class WorkflowDocumentoEmpleado
    {
        [MFWorkflow]
        [JsonConfEditor(Label = "Flujo de trabajo")]
        [DataMember]
        public MFIdentifier WorkflowValidacionesDocEmpleado { get; set; }

        [MFState]
        [DataMember]
        [JsonConfEditor(Label = "Documento Vigente")]
        public MFIdentifier EstadoDocumentoVigenteEmpleado { get; set; }
    }
}
