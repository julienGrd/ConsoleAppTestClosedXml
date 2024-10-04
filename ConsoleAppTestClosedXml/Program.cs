// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;

Console.WriteLine("Hello, World!");

CleanXlFile("StatdcK4-Trame_CAMSP_2023_VF_avec_fonctions-ORiGiNAL.xlsx", "StatdcK4-Trame_CAMSP_2023_VF_avec_fonctions-ORiGiNAL_clean.xlsx");


static void CleanXlFile(string pXlFullPath, string pDestFullPath = null)
{
    //-- CNSA CAMSP (2024) : rempl. formules par leur valeur sur la page AGR et suppr. 6 Feuilles temporaires 
    using (var lWorkbook = new XLWorkbook(pXlFullPath))
    {
        //-- si la page 'Enf_accomp_agr' existe : supprime toutes les formules et les remplace par leur valeur
        ReplaceAllFormulasWithActualValue(lWorkbook, "Enfants_accompagnes_AGR");

        //-- suppr les 5 feuilles intermediaires + Enfants_accompagnes
        DeleteSheets(lWorkbook, "Enfants_accompagnes",
                                                "feuille_intermediaire_file_acti",
                                                "feuille_intermediaire_present",
                                                "feuille_intermediaire_thera",
                                                "feuille_intermediaire_entres",
                                                "feuille_intermediaire_sortie");

        //-- Enregistre le fichier
        if (pDestFullPath == null)
            lWorkbook.Save();
        else
            lWorkbook.SaveAs(pDestFullPath);
    }
}


//internal static void UpdateDefinedNameValues(ref MemoryStream pContentStream, List<KeyValuePair<string, object>> pDefVars, out string pWarnMsg)
//{
//    using (var lWb = new XLWorkbook(pContentStream))
//    {
//        //-- récup. DefinedNames générales et celles de toutes les feuilles
//        //  à priori aucune méthode pour chercher 1 DefinedName ds tout le Xls
//        var lDefRanges = lWb.NamedRanges
//                        .Union(lWb.Worksheets.SelectMany(lWs => lWs.NamedRanges))
//                        .Where(lDn => pDefVars.Any(lDv => lDv.Key == lDn.Name)).ToList();
//        if (lDefRanges == null || lDefRanges.Count < pDefVars.Count)
//        {
//            //-- N'a pas trouvé ttes les var. à maj : retourne warning
//            var lNotFoundDn = (lDefRanges == null)
//                ? pDefVars.Select(lDv => lDv.Key).ToList()
//                : pDefVars.Select(lDv => lDv.Key).Except((lDefRanges).Select(lDr => lDr.Name)).ToList();
//            pWarnMsg = string.Format("{0} variable{1} non trouvée{1} : {2}", lNotFoundDn.Count, lNotFoundDn.Count == 1 ? "" : "s", string.Join(", ", lNotFoundDn));
//        }
//        else
//            pWarnMsg = null;

//        //int lValue = 1;
//        foreach (var lDefVar in pDefVars.Where(lDv => lDv.Value != null))
//        {
//            var lRange = lDefRanges.FirstOrDefault(lDn => lDn.Name == lDefVar.Key);
//            if (lRange != null)
//            {
//                var lCell = lWb.Cell(lRange.RefersTo);
//                if (lCell != null)
//                {
//                    if (lDefVar.Value is int)
//                        lCell.SetValue((int)lDefVar.Value);
//                    else if (lDefVar.Value is string)
//                        lCell.SetValue((string)lDefVar.Value);
//                    else if (lDefVar.Value is DateTime)
//                        lCell.SetValue((DateTime)lDefVar.Value);
//                }
//            }
//        }
//        lWb.SaveAs(pContentStream);
//    }
//}

// Supprime la formule de chaque cell de cette page en conservant la valeur de la cell
// (Excel mémorise pour chaque cell sa formule et sa valeur calculée)
static void ReplaceAllFormulasWithActualValue(XLWorkbook pWorkBook, string pSheetName)
{
    var lSheet = pWorkBook.Worksheets.SingleOrDefault(s => s.Name == pSheetName);
    if (lSheet != null)
    {
        //lSheet.RecalculateAllFormulas(); => 
        foreach (var lRow in lSheet.Rows())
        {
            foreach (IXLCell lCell in lRow.Cells())
            {
                if (lCell.HasFormula)
                {
                    //dynamic lValue = null;
                    //if(lCell.DataType == XLDataType.Number)
                    //    lValue = lCell.IsEmpty() ? 0 : lCell.Value.GetNumber();
                    //else if (lCell.DataType == XLDataType.Text)
                    //    lValue = lCell.IsEmpty() ? null : lCell.Value.GetText();
                    //else if (lCell.DataType == XLDataType.DateTime)
                    //    lValue = lCell.IsEmpty() ? null : lCell.Value.GetDateTime();
                    //lCell.Formula = null;

                    //semble suffire, a voir quand y aura des vrais valeurs dans le classeur
                    //var lValue = lCell.Value;
                    lCell.SetValue(lCell.CachedValue);
                }
            }
        }
    }
}

static void DeleteSheet(XLWorkbook pWorkBook, string pSheetName)
{
    if (pWorkBook.Worksheets.Any(w => w.Name == pSheetName))
        pWorkBook.Worksheets.Delete(pSheetName);
}

static void DeleteSheets(XLWorkbook pWorkBook, params string[] pSheetNames)
{
    foreach (var lSheet in pSheetNames)
        DeleteSheet(pWorkBook, lSheet);
}