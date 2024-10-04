using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AddOnCENCOSUD.App;
namespace AddOnCENCOSUD.DB_Structure
{
    class CreateStructure
    {
        public static void CreateStruct()
        {

            MDFields oMDFields = new MDFields();
            MDTables oMDTables = new MDTables();

            //oMDFields.CreateRegularField("OBPL", "EXX_ABPC_CEMP", "Codigo Empresa BPC", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50,
            // SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");

            //Tablas Nuevas 
            //oMDTables.CreateTableMD("EXX_AUPP_CNFG", "EXX - Configuración UpdPrice", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            //oMDTables.CreateTableMD("EXX_AUPP_UMLP", "EXX - Listas Unidad de Medida", SAPbobsCOM.BoUTBTableType.bott_NoObject);

            oMDTables.CreateTableMD("EXX_CCSD_TRANS", "EXX - CENCOSUD Transporte", SAPbobsCOM.BoUTBTableType.bott_Document);
            oMDTables.CreateTableMD("EXX_CCSD_TRAN1", "EXX - CENCOSUD Detalle", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

            #region TablaConfiguración
            //oMDFields.CreateRegularField("EXX_AUPP_CNFG", "EXX_AUPP_UPUM", "Actualizar unidad medidad?", SAPbobsCOM.BoFieldTypes.db_Alpha,
            //    SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO, null, null,
            //    new string[] { "N", "Y" },
            //    new string[] { "No", "Si" }, "N");
            //oMDFields.CreateRegularField("EXX_AUPP_CNFG", "EXX_AUPP_UOUM", "Unidad de Medida", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100,
            //   SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            //oMDFields.CreateRegularField("EXX_AUPP_UMLP", "EXX_AUPP_UMOC", "Código UM Origen", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
            //    SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            //oMDFields.CreateRegularField("EXX_AUPP_UMLP", "EXX_AUPP_UMOD", "Descripción UM Origen", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100,
            //    SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            //oMDFields.CreateRegularField("EXX_AUPP_UMLP", "EXX_AUPP_UMDC", "Código UM Destino", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
            //    SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            //oMDFields.CreateRegularField("EXX_AUPP_UMLP", "EXX_AUPP_UMDD", "Descripción UM Destino", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100,
            //    SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            #endregion

            #region CamposTabla1
            oMDFields.CreateRegularField("EXX_CCSD_TRANS", "EXX_CCSD_ARCH", "Archivo", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_DATE", "Fecha", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 20,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_GUIA", "N° Guia", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
                 SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_TRAN", "N° Trasnportista", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
                 SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_TNDA", "Tienda", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 12,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_CTON", "Carton", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Measurement, 0,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_CABO", "Cajas/Bolsas", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Measurement, 0,
                 SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_PLAS", "Plastico", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Measurement, 0,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_PLACA", "Placa", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 8,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_DESC", "Descarga", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_DIFT", "Diferencia Total", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_DIFC", "Diferencia Carton", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_DICB", "Diferencia Cajas/Bolsas", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("EXX_CCSD_TRAN1", "EXX_CCSD_DIFP", "Difenrecia Plastico", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            //oMDFields.CreateRegularField("EXX_AUPP_SPR1", "EXX_AUPP_PREC", "E-Commerce P.", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 0,
            //    SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");

            #endregion

            #region CamposOITM
            oMDFields.CreateRegularField("OITM", "EXX_CCSD_CLAR", "Clasificacion Articulo", SAPbobsCOM.BoFieldTypes.db_Alpha,
              SAPbobsCOM.BoFldSubTypes.st_None, 1, SAPbobsCOM.BoYesNoEnum.tNO, null, null,
              new string[] { "00", "01", "02", "03", "04" },
              new string[] { "Ninguno", "Carton", "Caja", "Bolsa", "Plastico" }, "0");

            oMDFields.CreateRegularField("OPDN", "EXX_CCSD_VALI", "Validación Addon", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
               SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("OPDN", "EXX_CCSD_VALD", "Validación por Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1,
                SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");
            oMDFields.CreateRegularField("PDN1", "EXX_CCSD_VALA", "Validación por articulo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20,
               SAPbobsCOM.BoYesNoEnum.tNO, null, null, null, null, "");

            #endregion

            if (!Globals.ExisteUDO("EXX_CCSD_TRANS"))
            {
                RegisterEXXLSPR();
            }

        }

        public static void RegisterEXXLSPR()
        {
            Globals.SBO_Application.ActivateMenuItem("4879");
            SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.ActiveForm;
            string title = oForm.Title;
            #region Step1
            oForm.Items.Item("4").Click();
            oForm.Items.Item("4").Click();
            SAPbouiCOM.EditText oEditCode = oForm.Items.Item("16").Specific;
            oEditCode.Value = "EXX_CCSD_TRANS";
            SAPbouiCOM.EditText oEditName = oForm.Items.Item("18").Specific;
            oEditName.Value = "EXX - CENCOSUD Transporte";
            SAPbouiCOM.ComboBox oComboType = oForm.Items.Item("20").Specific;
            oComboType.Select("3");
            SAPbouiCOM.EditText oEditTableName = oForm.Items.Item("62").Specific;
            oEditTableName.Value = "EXX_CCSD_TRANS";
            oForm.Items.Item("4").Click();
            #endregion
            #region Step2
            SAPbouiCOM.CheckBox oCheckFind = oForm.Items.Item("30").Specific;
            oCheckFind.Checked = true;
            SAPbouiCOM.CheckBox oCheckLog = oForm.Items.Item("35").Specific;
            oCheckLog.Checked = true;
            oForm.Items.Item("4").Click();
            #endregion
            #region Step3
            SAPbouiCOM.CheckBox oCheckDefForm = oForm.Items.Item("37").Specific;
            oCheckDefForm.Checked = true;
            SAPbouiCOM.OptionBtn oOptnNewForm = oForm.Items.Item("1250000092").Specific;
            oOptnNewForm.Selected = true;
            //SAPbouiCOM.CheckBox oCheckMenuItem = oForm.Items.Item("1250000084").Specific;
            //oCheckMenuItem.Checked = true;
            //SAPbouiCOM.EditText oEditMenuCpt = oForm.Items.Item("1250000085").Specific;
            //oEditMenuCpt.Value = "AsientosControl";
            //SAPbouiCOM.EditText oEditFather = oForm.Items.Item("1250000088").Specific;
            //oEditFather.Value = "43526";
            //SAPbouiCOM.EditText oEditPositon = oForm.Items.Item("1250000090").Specific;
            //oEditPositon.Value = "3";
            //SAPbouiCOM.EditText oEditUID = oForm.Items.Item("1250000094").Specific;
            //oEditUID.Value = "TPODOC";
            oForm.Items.Item("4").Click();
            #endregion
            #region Step4
            SAPbouiCOM.Matrix oMatrixFind = oForm.Items.Item("46").Specific;
            SAPbouiCOM.CheckBox oCheckMatFind;
            string[] ArrFind = { "U_EXX_CCSD_ARCH" };
            for (int j = 1; j < oMatrixFind.RowCount + 1; j++)
            {
                string value = oMatrixFind.Columns.Item("3").Cells.Item(j).Specific.Value;
                oCheckMatFind = oMatrixFind.Columns.Item("2").Cells.Item(j).Specific;
                if (ArrFind.Contains(value))
                    if (oCheckMatFind.Checked == false)
                        oMatrixFind.Columns.Item("2").Cells.Item(j).Click();
            }
            oForm.Items.Item("4").Click();
            #endregion
            #region Step5
            SAPbouiCOM.Matrix oMatrixDefault = oForm.Items.Item("42").Specific;
            SAPbouiCOM.CheckBox oCheckMatDefault;
            string[] ArrDefault = { "U_EXX_CCSD_ARCH" };
            for (int j = 1; j < oMatrixDefault.RowCount + 1; j++)
            {
                string value = oMatrixDefault.Columns.Item("3").Cells.Item(j).Specific.Value;
                oCheckMatDefault = oMatrixDefault.Columns.Item("2").Cells.Item(j).Specific;
                if (ArrDefault.Contains(value))
                    if (oCheckMatDefault.Checked == false)
                        oMatrixDefault.Columns.Item("2").Cells.Item(j).Click();
            }
            oForm.Items.Item("4").Click();
            #endregion
            #region Step6
            SAPbouiCOM.Matrix oMatrixSon = oForm.Items.Item("23").Specific;
            SAPbouiCOM.CheckBox oCheckMatSon;
            string[] ArrSon = { "EXX_CCSD_TRAN1" };
            for (int j = 1; j < oMatrixSon.RowCount + 1; j++)
            {
                string value = oMatrixSon.Columns.Item("1").Cells.Item(j).Specific.Value;
                oCheckMatSon = oMatrixSon.Columns.Item("2").Cells.Item(j).Specific;
                if (ArrSon.Contains(value))
                    if (oCheckMatSon.Checked == false)
                        oMatrixSon.Columns.Item("2").Cells.Item(j).Click();
            }
            oForm.Items.Item("4").Click();
            #endregion
            #region Step7
            SAPbouiCOM.ComboBox oComboSon = oForm.Items.Item("65").Specific;
            oComboSon.Select("1");
            SAPbouiCOM.Matrix oMatrixChildren = oForm.Items.Item("63").Specific;
            SAPbouiCOM.CheckBox oCheckChildren;
            //AGREGAR TODAS LAS COLUMNAS DE DETALLE
            string[] ArrChild1 = { "U_EXX_CCSD_DATE", "U_EXX_CCSD_GUIA", "U_EXX_CCSD_TNDA", "U_EXX_CCSD_CTON", "U_EXX_CCSD_CABO", 
                "U_EXX_CCSD_PLAS", "U_EXX_CCSD_PLACA", "U_EXX_CCSD_DESC", "U_EXX_CCSD_DIFT", "U_EXX_CCSD_DIFC", "U_EXX_CCSD_DICB", 
                "U_EXX_CCSD_DIFP" };
            for (int j = 1; j < oMatrixChildren.RowCount + 1; j++)
            {
                string value = oMatrixChildren.Columns.Item("3").Cells.Item(j).Specific.Value;
                oCheckChildren = oMatrixChildren.Columns.Item("2").Cells.Item(j).Specific;
                if (ArrChild1.Contains(value))
                    if (oCheckChildren.Checked == false)
                        oMatrixChildren.Columns.Item("2").Cells.Item(j).Click();
            }
            oForm.Items.Item("4").Click();
            oForm.Items.Item("5").Click();
            oForm.Items.Item("5").Click();
            #endregion
        }
    }
}
