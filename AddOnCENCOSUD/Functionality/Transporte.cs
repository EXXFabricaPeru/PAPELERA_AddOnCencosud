using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using AddOnCENCOSUD.App;
using AddOnCENCOSUD.DB_Structure;
using Microsoft.VisualBasic.Logging;
using SAPbobsCOM;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace AddOnCENCOSUD.Functionality
{
    public class Transporte
    {
        private static bool isLoadingItems = false;
        private static SAPbouiCOM.DataTable MyDataTable = null;
        private static string check = "N";
        private static SAPbouiCOM.Form oFormPadre = null;

        public static void Actions(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            SAPbouiCOM.ProgressBar pgrsBar = null;
            try
            {
                if (pVal.FormTypeEx == Globals.fmrUdoCENCOSUD)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.Action_Success)
                        {
                            if (pVal.ItemUID == "btnFile") OpenFile(oForm);
                            if (pVal.ItemUID == "btnCargar") CargarFile(oForm);
                            if (pVal.ItemUID == "btnValidar") ValidarFile(oForm);

                        }
                        if (pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "1") ActualizarGuia(oForm);
                        }
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.ItemUID == "1" && pVal.BeforeAction)
                        {
                            pgrsBar = Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                            GuardarRegistro(oForm);
                            pgrsBar.Stop();
                        }                       
                    }                                      
                }

                
                if (pVal.FormTypeEx == Globals.fmrUdoEstructuras)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.Action_Success)
                        {
                            if (pVal.ItemUID == "btnBF") CreateBF();
                        }
                    }
                }

                //if (pVal.FormTypeEx == "155")
                //{
                //    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                //    {
                //        if (pVal.BeforeAction)
                //        {
                //            if (pVal.ItemUID == "1") ValidarMaestroListas(oForm);
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                if (pgrsBar != null) pgrsBar.Stop();
                throw ex;
            }
        }

        public static void ActionsDataEvent(SAPbouiCOM.Form oForm, SAPbouiCOM.BusinessObjectInfo businessObjectInfo)
        {
            switch (businessObjectInfo.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                    if (oForm.TypeEx == "UDO_FT_EXX_AUPP_LSPR" && !businessObjectInfo.BeforeAction && businessObjectInfo.ActionSuccess)
                    {
                        SetEditableColumnsMatrix(oForm, false);
                    }
                    break;
            }
        }

        public static void LoadForm()
        {
            SAPbouiCOM.MenuItem menu = Globals.SBO_Application.Menus.Item("47616");
            SAPbouiCOM.OptionBtn oOptionBtn, oOptionBtn1 = null;
            SAPbouiCOM.Item oItem = null;

            try
            {
                if (menu.SubMenus.Count > 0)
                {
                    for (int i = 0; i < menu.SubMenus.Count; i++)
                    {
                        if (menu.SubMenus.Item(i).String.Contains("CCSD_TRANS"))
                        {
                            menu.SubMenus.Item(i).Activate();

                            var frmAux = (SAPbouiCOM.Form)Globals.SBO_Application.Forms.ActiveForm;
                            frmAux.AutoManaged = true;

                            ((SAPbouiCOM.EditText)frmAux.Items.Item("edtCntReg").Specific).Value = "0";

                            frmAux.DataSources.UserDataSources.Add("OpBtnDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                            oOptionBtn = (SAPbouiCOM.OptionBtn)frmAux.Items.Item("Item_5").Specific;
                            oOptionBtn.DataBind.SetBound(true, "", "OpBtnDS");
                            oOptionBtn1 = (SAPbouiCOM.OptionBtn)frmAux.Items.Item("Item_6").Specific;
                            oOptionBtn1.GroupWith("Item_5");
                            oOptionBtn1.DataBind.SetBound(true, "", "OpBtnDS");

                            //frmAux.Items.Item("Item_6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                            //frmAux.Items.Item("Item_6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True);

                            frmAux.Items.Item("btnValidar").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            frmAux.Items.Item("btnValidar").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True);

                            frmAux.Items.Item("btnFile").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            frmAux.Items.Item("btnFile").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True);

                            frmAux.Items.Item("0_U_G").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            frmAux.Items.Item("0_U_G").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True);

                            frmAux.Items.Item("btnCargar").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            frmAux.Items.Item("btnCargar").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True);

                            frmAux.Items.Item("edtCntReg").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            SetEditableColumnsMatrix(frmAux, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.SetStatusBarMessage("ERROR: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, true);
            }
        }

        public static void LoadFormEstr()
        {
            try
            {
                SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
                try
                {
                    oForm = Globals.SBO_Application.Forms.Item("FormEst_2");
                    Globals.SBO_Application.MessageBox("El formulario ya se encuentra abierto.");
                }
                catch //(Exception ex)
                {
                    SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
                    fcp = Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                    fcp.FormType = "FormEst";
                    fcp.UniqueID = "FormEst_2";
                    string FormName = "\\FormEst.srf";
                    fcp.XmlData = Globals.LoadFromXML(ref FormName);
                    oForm = Globals.SBO_Application.Forms.AddEx(fcp);
                }
                oForm.Top = 50;
                oForm.Left = 345;
                oForm.Visible = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void OpenFile(SAPbouiCOM.Form oForm)
        {
            try
            {
                Globals.OpenFileExcel(oForm, "20_U_E");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void CreateBF()
        {
            try
            {
                Globals.SBO_Application.SetStatusBarMessage("Creando BFs...", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                #region Let there be QCats UQ and FMS
                MDQueries oMDQueries = new MDQueries();

                #region QCats
                oMDQueries.CreateCategories("EXX_AddOn_CENCOSUD");
                #endregion

                #region UQ
                #region UQ HANA
                if (Globals.IsHana() == true)
                {
                    #region
                    //no identificado aun
                    oMDQueries.CreateQueries("EXX_AddOn_CENCOSUD", "BF_ObtenerGuia", AddOnCENCOSUD.Properties.Resources.BF_ObtenerGuia);
                    oMDQueries.CreateQueries("EXX_AddOn_CENCOSUD", "BF_ObtenerTransportista", AddOnCENCOSUD.Properties.Resources.BF_ObtenerTransportista);
                    #endregion
                }
                #endregion
                #region UQ SQL
                if (Globals.IsHana() == false)
                {
                    #region
                    oMDQueries.CreateQueries("EXX_AddOn_CENCOSUD", "BF_ObtenerGuia", AddOnCENCOSUD.Properties.Resources.BF_ObtenerGuia);
                    oMDQueries.CreateQueries("EXX_AddOn_CENCOSUD", "BF_ObtenerTransportista", AddOnCENCOSUD.Properties.Resources.BF_ObtenerTransportista);
                    #endregion

                }
                #endregion
                #endregion

                #region Old FMS

                //oMDQueries.RemoveFMS("BF_ObtenerMonedas", Globals.fmrUdoListaPrecios, "0_U_G", "C_0_1", "", "Y", "EXX_AddOn_UpdPrice");

                #endregion

                #region FMS
                oMDQueries.CreateFMS("BF_ObtenerGuia", Globals.fmrUdoCENCOSUD, "0_U_G", "C_0_2", "C_0_1", "N");
                oMDQueries.CreateFMS("BF_ObtenerTransportista", Globals.fmrUdoCENCOSUD, "0_U_G", "C_0_3", "", "N");

                #endregion

                #endregion

                Globals.SBO_Application.SetStatusBarMessage("BFs creadas correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void CargarFile(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.EditText oPath = oForm.Items.Item("20_U_E").Specific;

                if (oPath.Value.ToString() == "")
                {
                    throw new Exception("Debe ingresar un archivo para la carga.");
                }
                else
                {
                    LoadExcelFile(oPath.Value.ToString(), oForm);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void LoadExcelFile(string filename, SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("0_U_G").Specific;
                var lstCodItems = new List<string>();
                var lstDscError = new List<string>();
                var msjLog = string.Empty;
                List<DataTable> DataTables = new List<DataTable>();
                string line;
                string[] parts;
                int RecordLines = 0;
                var xmlMtx = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);

                isLoadingItems = true;

                DataTables = Globals.ImportExcel(filename);
                var lineNum = 1;
                //oForm.Freeze(true);

                var lastRowEmpty = string.IsNullOrWhiteSpace(oMatrix.GetCellSpecific("C_0_1", oMatrix.RowCount).Value)
                    || string.IsNullOrWhiteSpace(oMatrix.GetCellSpecific("C_0_2", oMatrix.RowCount).Value);

                foreach(DataRow row in DataTables[0].Rows)
                {
                    if (!lastRowEmpty) oMatrix.AddRow();
                    AgregarLineaMatrix(row, oMatrix, lineNum);
                    if (lastRowEmpty)
                    {
                        oMatrix.AddRow();
                        oMatrix.Columns.Item("C_0_1").Cells.Item(oMatrix.RowCount).Specific.Value = null;
                        oMatrix.Columns.Item("C_0_2").Cells.Item(oMatrix.RowCount).Specific.Value = null;
                    }
                    Globals.iRows++;
                    lineNum++;
                }
                    isLoadingItems = false;
                    //((SAPbouiCOM.EditText)oForm.Items.Item("edtCntReg").Specific).Value = (oMatrix.RowCount - 1).ToString();
                    msjLog = $"Proceso culminado \nRegistros cargados correctamente:{lineNum - 1}/{RecordLines - 2}";
                    /* lstDscError.ForEach(m =>
                     {
                         msjLog += "\n" + m;
                     });
                     */
                    MostrarCantidadRegistros(oForm);
                    //Globals.SBO_Application.MessageBox(msjLog);
                    Globals.SBO_Application.SetStatusBarMessage("Archivo cargado con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, false);               
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                isLoadingItems = false;
                oForm.Freeze(false);
            }

        }

        public static void AgregarLineaMatrix(DataRow parts, SAPbouiCOM.Matrix oMatrix, int iRows)
        {
            try
            {
                //if (oMatrix.Columns.Item("C_0_1").Cells.Item(iRows).Specific.Value.ToString() != "")
                //{
                //    iRows++;
                //}
                var maxMatrixLineNum = oMatrix.RowCount;
                //SAPbouiCOM.CheckBox oChkBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Col_0").Cells.Item(maxMatrixLineNum).Specific;
                //oChkBox.Checked = true;

                DateTime now = DateTime.Now;

                string formattedDate = string.Format("{0:yyyyMMdd}", parts[0]);
                //DateTime A = DateTime.ParseExact(parts[0].ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                oMatrix.Columns.Item("C_0_1").Cells.Item(maxMatrixLineNum).Specific.Value = formattedDate;
                oMatrix.Columns.Item("C_0_2").Cells.Item(maxMatrixLineNum).Specific.Value = parts[1].ToString() == "" ? "" : parts[1].ToString();
                oMatrix.Columns.Item("C_0_13").Cells.Item(maxMatrixLineNum).Specific.Value = parts[2].ToString() == "" ? "" : parts[2].ToString();
                oMatrix.Columns.Item("C_0_3").Cells.Item(maxMatrixLineNum).Specific.Value = parts[3].ToString() == "" ? "" : parts[3].ToString();
                oMatrix.Columns.Item("C_0_4").Cells.Item(maxMatrixLineNum).Specific.Value = double.Parse(parts[4].ToString() == "" ? "0" : parts[4].ToString());
                oMatrix.Columns.Item("C_0_5").Cells.Item(maxMatrixLineNum).Specific.Value = double.Parse(parts[5].ToString() == "" ? "0" : parts[5].ToString());
                oMatrix.Columns.Item("C_0_6").Cells.Item(maxMatrixLineNum).Specific.Value = double.Parse(parts[6].ToString() == "" ? "0" : parts[6].ToString());
                oMatrix.Columns.Item("C_0_7").Cells.Item(maxMatrixLineNum).Specific.Value = parts[7].ToString() == "" ? "" : parts[7].ToString();
                oMatrix.Columns.Item("C_0_8").Cells.Item(maxMatrixLineNum).Specific.Value = parts[8].ToString() == "" ? "" : parts[8].ToString();
                //oMatrix.Columns.Item("C_0_9").Cells.Item(maxMatrixLineNum).Specific.Value = parts[9].ToString();
                //oMatrix.Columns.Item("C_0_10").Cells.Item(maxMatrixLineNum).Specific.Value = parts[10].ToString();
                //oMatrix.Columns.Item("C_0_11").Cells.Item(maxMatrixLineNum).Specific.Value = parts[11].ToString();
                //oMatrix.Columns.Item("C_0_12").Cells.Item(maxMatrixLineNum).Specific.Value = parts[12].ToString();


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void MostrarCantidadRegistros(SAPbouiCOM.Form oForm)
        {
            var cntReg = 0;
            Task.Factory.StartNew(() =>
            {
                Thread.Sleep(200);
                var edtCntReg = (SAPbouiCOM.EditText)oForm.Items.Item("edtCntReg").Specific;
                var mtxItems = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;

                for (int i = 0; i < mtxItems.RowCount; i++)
                    if (!string.IsNullOrWhiteSpace(mtxItems.GetCellSpecific("C_0_1", i + 1).Value)) cntReg++;
                edtCntReg.Value = cntReg.ToString();
                //oForm.DataSources.DBDataSources.Item("@EXX_CCSD_TRANS").SetValue("U_EXX_CCSD_CNTREG", 0, cntReg.ToString());
            });
        }

        public static void OnClickAddMenu(SAPbouiCOM.Form oForm)
        {
            SetEditableColumnsMatrix(oForm, true);
        }

        private static void SetEditableColumnsMatrix(SAPbouiCOM.Form oForm, bool editable)
        {
            var matrix = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
            matrix.Columns.Item("C_0_1").Editable = editable;
            matrix.Columns.Item("C_0_2").Editable = editable;
            matrix.Columns.Item("C_0_3").Editable = editable;
            matrix.Columns.Item("C_0_4").Editable = editable;
            matrix.Columns.Item("C_0_5").Editable = editable;
            matrix.Columns.Item("C_0_6").Editable = editable;
            matrix.Columns.Item("C_0_7").Editable = editable;
            matrix.Columns.Item("C_0_8").Editable = editable;

        }
        public static void ValidarFile(SAPbouiCOM.Form oForm)
        {
            try {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("0_U_G").Specific;

                if (oForm.DataSources.UserDataSources.Item("OpBtnDS").Value == "2")
                {
                    ValidarxArticulo(oForm);
                }
                if (oForm.DataSources.UserDataSources.Item("OpBtnDS").Value == "1")
                {
                    ValidarTotal(oForm);
                }
                if (oForm.DataSources.UserDataSources.Item("OpBtnDS").Value != "2" && oForm.DataSources.UserDataSources.Item("OpBtnDS").Value != "1")
                {
                    string msjLog = "Debe seleccionar alguna de las opciones de tipo de validación";
                    Globals.SBO_Application.MessageBox(msjLog);

                }


            }
            catch (Exception ex) { }

        }


        public static void ValidarTotal(SAPbouiCOM.Form oForm)
        {
             SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("0_U_G").Specific;
            try
            {
                var maxMatrixLineNum = oMatrix.RowCount;

                for (int i = 1; oMatrix.RowCount > i; i++)
                {

                    string guia = oMatrix.Columns.Item("C_0_2").Cells.Item(i).Specific.Value.ToString();
                    string[] folio = guia.Split('-');
                    //double[] pesos = new double[3];
                    double  totalPesoSAP = 0;
                    double carton = double.Parse(oMatrix.Columns.Item("C_0_4").Cells.Item(i).Specific.Value.ToString());
                    double cajabolsa = double.Parse(oMatrix.Columns.Item("C_0_5").Cells.Item(i).Specific.Value.ToString());
                    double plastico = double.Parse(oMatrix.Columns.Item("C_0_6").Cells.Item(i).Specific.Value.ToString());
                    double totalPesoDoc = carton + cajabolsa+plastico;
                    //adaptar query para obtener los pesos de los tres articulos carton, cajas, plastico
                    Globals.Query = "SELECT T1.\"ItemCode\", T1.\"Quantity\", T2.\"U_EXX_CCSD_CLAR\", T0.\"U_EXX_CCSD_VALI\" FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                        "INNER JOIN OITM T2 ON T1.\"ItemCode\" = T2.\"ItemCode\" WHERE T0.\"FolioPref\" = '" + folio[0] + "' AND T0.\"FolioNum\" = '" + folio[1] + "'";
                    //T0.\"Series\" = '" + folio[0] + "'  AND
                    Globals.RunQuery(Globals.Query);
                    if (Globals.oRec.RecordCount > 0)
                    {
                        if (Globals.oRec.Fields.Item(3).Value.ToString() == "Y")
                        {
                            Globals.SBO_Application.SetStatusBarMessage("Documento: " + guia + " ya fue validado", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            continue;
                        }
                        while (!Globals.oRec.EoF)
                        {                           
                            string nombre = Globals.oRec.Fields.Item(2).Value.ToString();
                            double peso = Convert.ToInt32(Globals.oRec.Fields.Item(1).Value.ToString());
                            totalPesoSAP = totalPesoSAP + peso;
                            //comparar los de la matriz 

                            Globals.oRec.MoveNext();
                        }
                        if (totalPesoDoc > totalPesoSAP) oMatrix.Columns.Item("C_0_9").Cells.Item(i).Specific.Value = "Positivo";
                        if (totalPesoDoc < totalPesoSAP) oMatrix.Columns.Item("C_0_9").Cells.Item(i).Specific.Value = "Negativo";
                        if (totalPesoDoc == totalPesoSAP) oMatrix.Columns.Item("C_0_9").Cells.Item(i).Specific.Value = "Sin diferencia";
                    }

                    Globals.Release(Globals.oRec);



                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
        }


        public static void ValidarxArticulo(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("0_U_G").Specific;
                try
                {
                    var maxMatrixLineNum = oMatrix.RowCount;

                    for (int i = 1;oMatrix.RowCount >i; i++)
                    {

                        string guia = oMatrix.Columns.Item("C_0_2").Cells.Item(i).Specific.Value.ToString();
                        string[] folio = guia.Split('-');
                        //double[] pesos = new double[3];
                        double carton = double.Parse(oMatrix.Columns.Item("C_0_4").Cells.Item(i).Specific.Value.ToString());
                        double cajabolsa = double.Parse(oMatrix.Columns.Item("C_0_5").Cells.Item(i).Specific.Value.ToString());
                        double plastico = double.Parse(oMatrix.Columns.Item("C_0_6").Cells.Item(i).Specific.Value.ToString());
                        //adaptar query para obtener los pesos de los tres articulos carton, cajas, plastico
                        Globals.Query = "SELECT T1.\"ItemCode\", T1.\"Quantity\", T2.\"U_EXX_CCSD_CLAR\", T0.\"U_EXX_CCSD_VALI\" FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                        "INNER JOIN OITM T2 ON T1.\"ItemCode\" = T2.\"ItemCode\" WHERE T0.\"FolioPref\" = '" + folio[0] + "' AND T0.\"FolioNum\" = '" + folio[1] + "'";
                        //T0.\"Series\" = '" + folio[0] + "'  AND 
                        Globals.RunQuery(Globals.Query);
                        if(Globals.oRec.RecordCount>0)
                        {
                            for (int a = 0; Globals.oRec.RecordCount > a; a++)
                            {
                                if(Globals.oRec.EoF)
                                    continue;

                                if (Globals.oRec.Fields.Item(3).Value.ToString() == "Y")
                                {
                                    Globals.SBO_Application.SetStatusBarMessage("Documento: " + guia + " ya fue validado", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                    continue;
                                }
                                string nombre = Globals.oRec.Fields.Item(2).Value.ToString();
                                double peso = Convert.ToInt32(Globals.oRec.Fields.Item(1).Value.ToString());

                                //comparar los de la matriz 
                                //campo de usuario en maestro de articulos con opciones desplegables 00 - Ninguno, 01 - Carton, etc
                                //validar que si o si tenga apeso sino culminar validacion de esa fila
                                if (nombre == "01")
                                {
                                    if (carton > peso) oMatrix.Columns.Item("C_0_10").Cells.Item(i).Specific.Value = "Positivo";
                                    if (carton < peso) oMatrix.Columns.Item("C_0_10").Cells.Item(i).Specific.Value = "Negativo";
                                    if (carton == peso) oMatrix.Columns.Item("C_0_10").Cells.Item(i).Specific.Value = "Sin diferencia";

                                }
                                if (nombre == "02" || nombre == "03")
                                {
                                    if (cajabolsa > peso) oMatrix.Columns.Item("C_0_11").Cells.Item(i).Specific.Value = "Positivo";
                                    if (cajabolsa < peso) oMatrix.Columns.Item("C_0_11").Cells.Item(i).Specific.Value = "Negativo";
                                    if (cajabolsa == peso) oMatrix.Columns.Item("C_0_11").Cells.Item(i).Specific.Value = "Sin diferencia";

                                }
                                if (nombre == "04")
                                {
                                    if (plastico > peso) oMatrix.Columns.Item("C_0_12").Cells.Item(i).Specific.Value = "Positivo";
                                    if (plastico < peso) oMatrix.Columns.Item("C_0_12").Cells.Item(i).Specific.Value = "Negativo";
                                    if (plastico == peso) oMatrix.Columns.Item("C_0_12").Cells.Item(i).Specific.Value = "Sin diferencia";

                                }
                                Globals.oRec.MoveNext();
                            }
                        }

                        Globals.Release(Globals.oRec);


                       
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message.ToString());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ActualizarGuia(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.Documents oDoc = Globals.oCompany.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes);
            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("0_U_G").Specific;
                for (int i = 1; oMatrix.RowCount >i; i++)
                {
                    string guia = oMatrix.Columns.Item("C_0_2").Cells.Item(i).Specific.Value.ToString();
                    string valorDoc = oMatrix.Columns.Item("C_0_9").Cells.Item(i).Specific.Value.ToString();
                    string valorPlastico = oMatrix.Columns.Item("C_0_12").Cells.Item(i).Specific.Value.ToString();
                    string valorCajaBolsa = oMatrix.Columns.Item("C_0_11").Cells.Item(i).Specific.Value.ToString();
                    string valorCarton = oMatrix.Columns.Item("C_0_10").Cells.Item(i).Specific.Value.ToString();
                    int DocEntry = GetDocEntry(guia);
                    if (!oDoc.GetByKey(DocEntry))
                        throw new Exception($"La guía {guia} no se encuentra registrado en la sociedad");
                    oDoc.UserFields.Fields.Item("U_EXX_CCSD_VALI").Value = "Y";
                    oDoc.UserFields.Fields.Item("U_EXX_TIPOOPER").Value = "02";
                    Globals.Query = "SELECT T2.\"U_EXX_CCSD_CLAR\", T0.\"U_EXX_CCSD_VALI\" FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                            "INNER JOIN OITM T2 ON T1.\"ItemCode\" = T2.\"ItemCode\" WHERE T2.\"ItemCode\" = '" + oDoc.Lines.ItemCode + "'";
                    Globals.RunQuery(Globals.Query);
                    if (Globals.oRec.Fields.Item(1).Value.ToString() == "Y")
                    {
                        Globals.SBO_Application.SetStatusBarMessage("Documento: " + guia + " ya fue validado y actualizado", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        continue;
                    }
                    if (oForm.DataSources.UserDataSources.Item("OpBtnDS").Value == "1")
                    {
                        if (!string.IsNullOrEmpty(valorDoc)) oDoc.UserFields.Fields.Item("U_EXX_CCSD_VALD").Value = valorDoc;
                    }
                    if (oForm.DataSources.UserDataSources.Item("OpBtnDS").Value == "2")
                    {
                        for (int a = 0; a < oDoc.Lines.Count; a++)
                        {
                            oDoc.Lines.SetCurrentLine(a);
                            

                            string nombre = Globals.oRec.Fields.Item(0).Value.ToString();
                            if (nombre == "01")
                            {
                                if (!string.IsNullOrEmpty(valorCarton)) oDoc.Lines.UserFields.Fields.Item("U_EXX_CCSD_VALA").Value = valorCarton;
                            }
                            if (nombre == "02" || nombre == "03")
                            {
                                if (!string.IsNullOrEmpty(valorCajaBolsa)) oDoc.Lines.UserFields.Fields.Item("U_EXX_CCSD_VALA").Value = valorCajaBolsa;
                            }
                            if (nombre == "04")
                            {
                                if (!string.IsNullOrEmpty(valorPlastico)) oDoc.Lines.UserFields.Fields.Item("U_EXX_CCSD_VALA").Value = valorPlastico;
                            }

                        }
                    }
                    Globals.Release(Globals.oRec);

                    if (oDoc.Update() != 0)
                        throw new Exception(Globals.oCompany.GetLastErrorDescription());                    
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static int GetDocEntry(String guia)
        {
            try
            {
                string sDocNum = "";
                int Num = 0;
                string[] folio = guia.Split('-');
                Globals.Query = "SELECT \"DocEntry\" FROM OPDN T0 WHERE T0.\"FolioPref\" = '" + folio[0] + "' AND T0.\"FolioNum\" = '" + folio[1] + "'";
                Globals.oRec = Globals.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                Globals.oRec.DoQuery(Globals.Query);
                if (Globals.oRec.RecordCount > 0)
                {
                    sDocNum = Globals.oRec.Fields.Item(0).Value.ToString();
                }
                Globals.Release(Globals.oRec);

                if (sDocNum == "")
                {
                    throw new Exception("Error al obtener DocNum");
                }
                Num = Convert.ToInt32(sDocNum);
                return Num;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void GuardarRegistro(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("0_U_G").Specific;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
