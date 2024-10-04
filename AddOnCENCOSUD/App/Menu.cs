using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnCENCOSUD.App
{
    public static class Menu
    {

        public static void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);
            SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);
            oMenus = Globals.SBO_Application.Menus;
            SAPbouiCOM.MenuCreationParams oCreationPackage = default(SAPbouiCOM.MenuCreationParams);
            oCreationPackage = Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
            try
            {
                oMenuItem = Globals.SBO_Application.Menus.Item("2304");
                oMenus = oMenuItem.SubMenus;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "EXX_CCSD";
                oCreationPackage.String = "EXX - Transporte CENCOSUD";
                oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;

                //oCreationPackage.Image = ""//ruta iamgen

                if (!(oMenus.Exists("EXX_CCSD")))
                {
                    oMenus.AddEx(oCreationPackage);
                }

                oMenuItem = Globals.SBO_Application.Menus.Item("EXX_CCSD");
                oMenus = oMenuItem.SubMenus;

                //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //oCreationPackage.UniqueID = "EXX_CCSD0";
                //oCreationPackage.String = "EXX - Estructuras";
                //oMenus.AddEx(oCreationPackage);

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "EXX_CCSD1";
                oCreationPackage.String = "EXX - Transporte CENCOSUD";
                oMenus.AddEx(oCreationPackage);

            }
            catch (Exception ex)
            {
                Globals.SBO_Application.SetStatusBarMessage(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
        }
    }
}
