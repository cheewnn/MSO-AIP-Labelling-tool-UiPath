using System;
using System.IO;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using SNT.OfficeLabelTool.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;
using Microsoft.Office;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.Collections.Generic;

namespace SNT.OfficeLabelTool.Activities
{
    [LocalizedDisplayName(nameof(Resources.GetSensitivityLabel_DisplayName))]
    [LocalizedDescription(nameof(Resources.GetSensitivityLabel_Description))]
    public class GetSensitivityLabel : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetSensitivityLabel_FilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetSensitivityLabel_FilePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> FilePath { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetSensitivityLabel_LabelId_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetSensitivityLabel_LabelId_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> LabelId { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetSensitivityLabel_LabelName_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetSensitivityLabel_LabelName_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> LabelName { get; set; }

        [LocalizedDisplayName(nameof(Resources.GetSensitivityLabel_SiteId_DisplayName))]
        [LocalizedDescription(nameof(Resources.GetSensitivityLabel_SiteId_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> SiteId { get; set; }

        #endregion


        #region Constructors

        public GetSensitivityLabel()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (FilePath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(FilePath)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = FilePath.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////
            Excel.Application oXL = null;
            Excel.Workbook oWorkBook;
            Word.Application oW = null;
            Word.Document oDocument;
            PowerPoint.Application oPPT = null;
            PowerPoint.Presentation oPresentation;
            Microsoft.Office.Core.LabelInfo o_LabelInfo;
            string labelid = null;
            string siteid = null;
            string labelname = null;
            try
            {
                if (Path.GetExtension(FilePath.Get(context)).Contains(".xls"))
                {
                    System.Console.WriteLine("Excel Application");
                    oXL = new Excel.Application { Visible = false };
                    oWorkBook = (Excel.Workbook)(oXL.Workbooks.Open(filepath));
                    System.Console.WriteLine("Getting label");
                    o_LabelInfo = oWorkBook.SensitivityLabel.GetLabel();
                    System.Console.WriteLine("Label ID: " + o_LabelInfo.LabelId);
                    labelid = o_LabelInfo.LabelId;
                    labelname = o_LabelInfo.LabelName;
                    siteid = o_LabelInfo.SiteId;
                    oWorkBook.Application.Quit();
                    Marshal.ReleaseComObject(oWorkBook);
                    Marshal.ReleaseComObject(oXL);

                }
                else if (Path.GetExtension(FilePath.Get(context)).Contains(".doc"))
                {
                    System.Console.WriteLine("Word Application");
                    oW = new Word.Application { Visible = false };
                    oDocument = (Word.Document)(oW.Documents.Open(filepath));
                    System.Console.WriteLine("Getting label");
                    o_LabelInfo = oDocument.SensitivityLabel.GetLabel();
                    System.Console.WriteLine("Label ID: " + o_LabelInfo.LabelId);
                    labelid = o_LabelInfo.LabelId;
                    labelname = o_LabelInfo.LabelName;
                    siteid = o_LabelInfo.SiteId;
                    oDocument.Application.Quit();
                    Marshal.ReleaseComObject(oDocument);
                    Marshal.ReleaseComObject(oW);

                }
                else if (Path.GetExtension(FilePath.Get(context)).Contains(".ppt"))
                {
                    System.Console.WriteLine("Powerpoint Application");
                    oPPT = new PowerPoint.Application { Visible = MsoTriState.msoFalse };
                    oPresentation = (PowerPoint.Presentation)(oPPT.Presentations.Open(filepath));
                    System.Console.WriteLine("Getting label");
                    o_LabelInfo = oPresentation.SensitivityLabel.GetLabel();
                    System.Console.WriteLine("Label ID: " + o_LabelInfo.LabelId);
                    labelid = o_LabelInfo.LabelId;
                    labelname = o_LabelInfo.LabelName;
                    siteid = o_LabelInfo.SiteId;
                    oPresentation.Application.Quit();
                    Marshal.ReleaseComObject(oPresentation);
                    Marshal.ReleaseComObject(oPPT);

                }
                else
                {
                    labelid = "";
                    labelname = "";
                    siteid = "";
                    System.Console.WriteLine("Invalid file.");
                    throw new Exception("File path is invalid or is not an MSO office file.");
                }
            }

            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);

                List<dynamic> officeApplications = new List<dynamic> { oXL, oW, oPPT };
                foreach (var officeApp in officeApplications)
                {
                    if (officeApp != null)
                    {
                        try
                        {
                            //Attempt to close application
                            officeApp.Quit();
                        }
                        catch (Exception cleanupEx)
                        {
                            System.Console.WriteLine("Exception cleanup error: " + cleanupEx.Message);
                        }
                        finally
                        {
                            // Release COM Object
                            Marshal.ReleaseComObject(officeApp);
                        }
                    }
                }
                // Rethrow exception to UiPath workflow
                throw new Exception(ex.Message);
            }

            // Outputs
            return (ctx) => {
                LabelId.Set(ctx, labelid);
                LabelName.Set(ctx, labelname);
                SiteId.Set(ctx, siteid);
            };
        }

        #endregion
    }
}

