using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kompas6API5;
using Kompas6API7;
using Kompas6Constants;
using KAPITypes;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace ConsoleApp_Kompas
{
    class Program
    {
        static void Main(string[] args)
        {
            _Activation actv = new _Activation();
            actv.act();
        }
    }
    class _Activation
    {
        public KompasObject kompas;
        public ksDocument2D dock;
        public ksDocumentParam documentParam;
        //public ksSheetOptions sheetOptions; - для старых версии
        public ksSheetPar sheetPar;
        public ksStandartSheet standartSheet;
        public string str;
        public int SHEET_OPTIONS_EX = 4;

        public void act()
        {
            Type c = Type.GetTypeFromProgID("KOMPAS.Application.5");
            kompas = (KompasObject)Activator.CreateInstance(c);
            documentParam = (ksDocumentParam)kompas.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
            
            documentParam.type = (int)DocType.lt_DocSheetStandart;
            documentParam.regime = 0;
            sheetPar = (ksSheetPar)documentParam.GetLayoutParam();
            str = kompas.ksSystemPath(0) + @"\graphic.lyt";
            sheetPar.layoutName = str;
            sheetPar.shtType = 1;
            standartSheet = (ksStandartSheet)sheetPar.GetSheetParam();
            standartSheet.direct = true;
            standartSheet.format = 3;
            standartSheet.multiply = 1;
            dock = (ksDocument2D)kompas.Document2D();
            dock.ksCreateDocument(documentParam);

            //dock = (ksDocument2D)kompas.Document2D();                                                             // ДЛЯ СТАРОГО КОМПАСА
            //dock.ksCreateDocument(documentParam);                                                                 // ДЛЯ СТАРОГО КОМПАСА
            //sheetOptions = (ksSheetOptions)kompas.GetParamStruct((short)StructType2DEnum.ko_SheetOptions);        // ДЛЯ СТАРОГО КОМПАСА
            //dock.ksGetDocOptions(SHEET_OPTIONS_EX, sheetOptions);                                                 // ДЛЯ СТАРОГО КОМПАСА
            //sheetOptions.sheetType = false;                                                                       // ДЛЯ СТАРОГО КОМПАСА
            //standartSheet = (ksStandartSheet)sheetOptions.GetSheetParam(false);                                   // ДЛЯ СТАРОГО КОМПАСА
            //standartSheet.format = 3;                                                                             // ДЛЯ СТАРОГО КОМПАСА
            //standartSheet.direct = true;                                                                          // ДЛЯ СТАРОГО КОМПАСА
            //sheetOptions.sheetType = true;                                                                        // ДЛЯ СТАРОГО КОМПАСА
            //dock.ksSetDocOptions(SHEET_OPTIONS_EX, sheetOptions);                                                 // ДЛЯ СТАРОГО КОМПАСА
            kompas.Visible = true;
                
        }

        //public void dokiact()
        //{

        //    docpar.Init();
        //    docpar.type = (int)DocType.lt_DocSheetStandart;
        //    dock.ksCreateDocument(docpar);
        //}           
    }
}
