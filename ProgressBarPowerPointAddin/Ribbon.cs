using Microsoft.Office.Core;
using ProgressBarPowerPointAddin.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

// TODO:  リボン (XML) アイテムを有効にするには、次の手順に従います:

// 1: 次のコード ブロックを ThisAddin、ThisWorkbook、ThisDocument のいずれかのクラスにコピーします。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. ボタンのクリックなど、ユーザーの操作を処理するためのコールバック メソッドを、このクラスの
//    "リボンのコールバック" 領域に作成します。メモ: このリボンがリボン デザイナーからエクスポートされたものである場合は、
//    イベント ハンドラー内のコードをコールバック メソッドに移動し、リボン拡張機能 (RibbonX) のプログラミング モデルで
//    動作するように、コードを変更します。

// 3. リボン XML ファイルのコントロール タグに、コードで適切なコールバック メソッドを識別するための属性を割り当てます。  

// 詳細については、Visual Studio Tools for Office ヘルプにあるリボン XML のドキュメントを参照してください。


namespace ProgressBarPowerPointAddin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ProgressBarPowerPointAddin.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック
        //ここでコールバック メソッドを作成します。コールバック メソッドの追加について詳しくは https://go.microsoft.com/fwlink/?LinkID=271226 をご覧ください

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnButtonShapeProgressButton(IRibbonControl control)
        {
            //var currentSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides[Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex];
            //currentSlide.cha
            var app = Globals.ThisAddIn.Application.ActiveWindow;
            if (checkSelectionCount() == false)
                return;
            int forecolor = 0x009c7141;
            int bgcolor = 0x00eed7bd;
            UserInputForm form = new UserInputForm();
            if (form.ShowDialog() == DialogResult.Cancel)
                return;
            float value = form.Value * 0.01f;

            var selection = app.Selection.ShapeRange;
            //form.SelectedObject = selection.Cast<Shape>().First();
            foreach (PPT.Shape shp in selection)
            {
                if (isShapeSupported(shp) == false)
                    continue;
                if (shp.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillMixed || shp.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillSolid)
                {
                    if (shp.Line.ForeColor.RGB != 0x00ffffff)
                        forecolor = shp.Line.ForeColor.RGB;
                    if (shp.Fill.ForeColor.RGB != 0x00ffffff)
                        bgcolor = shp.Fill.ForeColor.RGB;
                }
                else if (shp.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillGradient)//既に設定している場合
                {
                    var grads = shp.Fill.GradientStops.Cast<Microsoft.Office.Core.GradientStop>();
                    //forecolor = grads.First().Color.RGB;
                    if (shp.Line.ForeColor.RGB != 0x00ffffff)
                        forecolor = shp.Line.ForeColor.RGB;
                    bgcolor = grads.Last().Color.RGB;
                }
                else
                    return;
                if (value == 1.0)
                {
                    shp.Fill.TwoColorGradient(Microsoft.Office.Core.MsoGradientStyle.msoGradientVertical, 1);
                    shp.Fill.GradientStops.Insert(forecolor, 1.0f, Index: 1);
                    //shp.Fill.GradientStops.Insert(forecolor, value, Index: 2);
                    //shp.Fill.GradientStops.Insert(bgcolor, value + 0.0001f, Index: 3);
                    shp.Fill.GradientStops.Insert(bgcolor, 1.0f, Index: 4);
                }
                shp.Fill.TwoColorGradient(Microsoft.Office.Core.MsoGradientStyle.msoGradientVertical, 1);
                shp.Fill.GradientStops.Insert(forecolor, 0, Index: 1);
                shp.Fill.GradientStops.Insert(forecolor, value, Index: 2);
                shp.Fill.GradientStops.Insert(bgcolor, value + 0.0001f, Index: 3);
                shp.Fill.GradientStops.Insert(bgcolor, 1.0f, Index: 4);
                while (shp.Fill.GradientStops.Count > 4)
                    shp.Fill.GradientStops.Delete();
            }
        }

        public void SetSupportedByCondition()// not used yet
        {
            ribbon.InvalidateControl("buttonShapeProgress");
        }

        private bool buttonShapeProgress_getEnabled(IRibbonControl control)//not used yet
        {
            var app = Globals.ThisAddIn.Application.ActiveWindow;
            var selection = app.Selection.ShapeRange as PPT.ShapeRange;
            if (selection.Count == 0)
                return false;
            bool supported = false;
            foreach (PPT.Shape shape in selection)
            {
                supported |= isShapeSupported(shape);
            }
            return supported;
        }
        private bool isShapeSupported(PPT.Shape shape)
        {
            switch (shape.Type)
            {
                case MsoShapeType.msoAutoShape:
                case MsoShapeType.msoPlaceholder:
                case MsoShapeType.msoPicture:
                case MsoShapeType.msoLinkedPicture:
                case MsoShapeType.msoCanvas:
                case MsoShapeType.msoTextBox:
                    return true;
                default:
                    return false;
            }
        }
        public Bitmap buttonShapeProgress_GetImage(IRibbonControl control)
        {
            return Resources.TagIcon;
        }
        #endregion
        bool checkSelectionCount()
        {
            var app = Globals.ThisAddIn.Application.ActiveWindow;
            int count = 0;
            try
            {
                count = app.Selection.ShapeRange.Count;
            }
            catch (Exception)
            {
                MessageBox.Show("Shapeを選択してください");
            }

            return count > 0;

        }

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
