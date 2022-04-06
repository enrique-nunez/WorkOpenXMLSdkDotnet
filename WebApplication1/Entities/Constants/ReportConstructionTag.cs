namespace WebApplication1.Entities.Constants
{
    public class ReportConstructionTag
    {
        /// <summary>
        /// Tag Table Inspection
        /// </summary>
        public const string REPORT_CONSTRUCTION_INSPECTION_DES_RAZONSOCIAL = "${{Inspeccion:des_razonsocial}}";
        public const string REPORT_CONSTRUCTION_INSPECTION_DES_DIRECCION = "${{Inspeccion:des_direccion}}";
        public const string REPORT_CONSTRUCTION_INSPECTION_FECHA_INSPECCION = "${{Inspeccion: fechaInspeccion}}";
        public const string REPORT_CONSTRUCTION_INSPECTION_NOMBRE_USUARIO_INSPECTOR = "$Nombre del Inspector {{Inspeccion: nombreUsuarioInspector}}";
        public const string REPORT_CONSTRUCTION_INSPECTION_NOMBRE_CONTRATANTE_DES_RAZONSOCIAL = "$Nombre de la empresa contratante del seguro {{Inspeccion:des_razonsocial}}";
        public const string REPORT_CONSTRUCTION_INSPECTION_NUM_LATITUD = "$En formato decimal {{Inspeccion:num_latitud}} ";
        public const string REPORT_CONSTRUCTION_INSPECTION_NUM_LONGITUD = "$En formato decimal {{Inspeccion:num_longitud}}";

        /// <summary>
        /// Tag Table Interviewed
        /// </summary>
        public const string REPORT_CONSTRUCTION_INTERVIEWED_NOM_ENTREVISTADO = "${{Entrevistado:nom_entrevistado}}";
        public const string REPORT_CONSTRUCTION_INTERVIEWED_NOM_CARGO = "${{Entrevistado:nom_cargo}}";
        public const string REPORT_CONSTRUCTION_INTERVIEWED_CORREO = "${{Entrevistado:correo}}";
        public const string REPORT_CONSTRUCTION_INTERVIEWED_TELEFONO = "${{Entrevistado:telefono}}";

        /// <summary>
        /// Tag Table Risk
        /// </summary>
        public const string REPORT_CONSTRUCTION_RISK_DES_TERREMOTO = "${{Riesgo:des_terremoto}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_INUNDACION = "${{Riesgo:des_inundacion}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_LLUVIA = "${{Riesgo:des_lluvia}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_MAREMOTO = "${{Riesgo:des_maremoto}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_GEOLOGICO = "${{Riesgo:des_geologico}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_GEOTECNIC = "${{Riesgo:des_geotecnic}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_EXPOSIC = "${{Riesgo:des_exposic}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_INCENDIO = "${{Riesgo:des_incendio}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_PERD_BEN = "${{Riesgo:des_perd_ben}}";
        public const string REPORT_CONSTRUCTION_RISK_DES_RESP_CIVIL = "${{Resgo:des_resp_civil}}";

        /// <summary>
        /// Tag Table Sinister
        /// </summary>
        public const string REPORT_CONSTRUCTION_SINISTER_FEC_SINIESTRO = "${{Siniestro:fec_siniestro}}";
        public const string REPORT_CONSTRUCTION_SINISTER_DES_SINIESTRO = "${{Siniestro:des_siniestro}}";
        public const string REPORT_CONSTRUCTION_SINISTER_DES_TIE_PARAL = "${{Siniestro:des_tie_paral}}";

        /// <summary>
        /// Tag Table Link
        /// </summary>
        public const string REPORT_CONSTRUCTION_LINK_DES_ENLACE = "${{Enlace: des_enlace}}";

        /// <summary>
        /// Tag Table Archive
        /// </summary>
        public const string REPORT_CONSTRUCTION_ARCHIVE_LOGO = "${{Logo_tipo01}}";
        public const string REPORT_CONSTRUCTION_ARCHIVE_FIRMA = "${{firma_tipo02}}";
    }
}
