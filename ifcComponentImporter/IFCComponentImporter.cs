using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.IFC;
using Autodesk.Revit.DB.Analysis;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

[Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
public class IFCComponentImporter : IExternalCommand
{
    [DllImport("USER32.DLL", SetLastError = true)]
    static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr PostMessage(IntPtr hWnd, uint msg, uint wParam, uint lParam);

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("USER32.DLL")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

    public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

    /// <summary>
    /// Callback method to be used when enumerating windows.
    /// </summary>
    /// <param name="handle">Handle of the next window</param>
    /// <param name="pointer">Pointer to a GCHandle that holds a reference to the list to fill</param>
    /// <returns>True to continue the enumeration, false to bail</returns>
    private static bool EnumWindow(IntPtr handle, IntPtr pointer)
    {
        GCHandle gch = GCHandle.FromIntPtr(pointer);
        List<IntPtr> list = gch.Target as List<IntPtr>;
        if (list == null)
        {
            throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
        }
        list.Add(handle);
        //  You can modify this to check to see if you want to cancel the operation, then return a null here
        return true;
    }


    public static List<IntPtr> GetChildWindows(IntPtr parent)
    {
        List<IntPtr> result = new List<IntPtr>();
        GCHandle listHandle = GCHandle.Alloc(result);
        try
        {
            EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
            EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
        }
        finally
        {
            if (listHandle.IsAllocated)
                listHandle.Free();
        }
        return result;
    }
    public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            string tmpPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\ifcComponentImporter";
            string filePath = "";
            string rvtPath = "";
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "IFC Files|*.ifc";
            openFileDialog1.Title = "Select a ifc File";
            if(!Directory.Exists(tmpPath))
            {
                Directory.CreateDirectory(tmpPath);
            }
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
            }
            Result result;
            var ifcops = new IFCImportOptions();
            Document document = commandData.Application.Application.OpenIFCDocument(filePath, ifcops);

            DocumentPreviewSettings settings = document.GetDocumentPreviewSettings();
            FilteredElementCollector collector = new FilteredElementCollector(document);
            collector.OfClass(typeof(ViewPlan));
            FilteredElementCollector uicollector = new FilteredElementCollector(commandData.Application.ActiveUIDocument.Document);
            uicollector.OfClass(typeof(ViewPlan));
            Func<ViewPlan, bool> isValidForPreview = v => settings.IsViewIdValidForPreview(v.Id);

            ViewPlan viewForPreview = collector.OfType<ViewPlan>().First<ViewPlan>(isValidForPreview);
            using (Transaction tx = new Transaction(document))
            {
                tx.Start("tr");
                viewForPreview.SetWorksetVisibility(viewForPreview.WorksetId, WorksetVisibility.Visible);
                tx.Commit();
            }
            SaveAsOptions options = new SaveAsOptions();
            options.PreviewViewId = viewForPreview.Id;

            rvtPath = tmpPath +@"\"+ Path.ChangeExtension(Path.GetFileName(filePath), "rvt");
         
            if (File.Exists(rvtPath))
            {
                File.Delete(rvtPath);
            }
            document.SaveAs(rvtPath, options);

            Process process = new Process();
            process.StartInfo.FileName="PostAction.exe";
            process.StartInfo.Arguments = rvtPath;
            process.Start();

               
            return 0;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Autodesk.Revit.UI.Result.Failed;
        }
    }

}
