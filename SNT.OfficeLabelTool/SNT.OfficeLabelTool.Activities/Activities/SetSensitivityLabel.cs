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
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace SNT.OfficeLabelTool.Activities
{
    [LocalizedDisplayName(nameof(Resources.SetSensitivityLabel_DisplayName))]
    [LocalizedDescription(nameof(Resources.SetSensitivityLabel_Description))]
    public class SetSensitivityLabel : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedDisplayName(nameof(Resources.SetSensitivityLabel_FilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.SetSensitivityLabel_FilePath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> FilePath { get; set; }

        [LocalizedDisplayName(nameof(Resources.SetSensitivityLabel_LabelId_DisplayName))]
        [LocalizedDescription(nameof(Resources.SetSensitivityLabel_LabelId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> LabelId { get; set; }

        [LocalizedDisplayName(nameof(Resources.SetSensitivityLabel_LabelName_DisplayName))]
        [LocalizedDescription(nameof(Resources.SetSensitivityLabel_LabelName_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> LabelName { get; set; }

        [LocalizedDisplayName(nameof(Resources.SetSensitivityLabel_SiteId_DisplayName))]
        [LocalizedDescription(nameof(Resources.SetSensitivityLabel_SiteId_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> SiteId { get; set; }

        [LocalizedDisplayName(nameof(Resources.SetSensitivityLabel_Result_DisplayName))]
        [LocalizedDescription(nameof(Resources.SetSensitivityLabel_Result_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<bool> Result { get; set; }

        #endregion


        #region Constructors

        public SetSensitivityLabel()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {
            if (FilePath == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(FilePath)));
            if (LabelId == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(LabelId)));
            if (LabelName == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(LabelName)));
            if (SiteId == null) metadata.AddValidationError(string.Format(Resources.ValidationValue_Error, nameof(SiteId)));

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var filepath = FilePath.Get(context);
            var labelid = LabelId.Get(context);
            var labelname = LabelName.Get(context);
            var siteid = SiteId.Get(context);

            ///////////////////////////
            // Add execution logic HERE
            ///////////////////////////
            Excel.Application oXL=null;
            Excel.Workbook oWorkBook;
            Word.Application oW=null;
            Word.Document oDocument;
            PowerPoint.Application oPPT = null;
            PowerPoint.Presentation oPresentation = null;
            Microsoft.Office.Core.LabelInfo o_LabelInfo;
            bool result = false;

            try
            {
                if (Path.GetExtension(filepath).Contains(".xls"))
                {
                    System.Console.WriteLine("Excel Application");
                    oXL = new Excel.Application { Visible = false };
                    oWorkBook = (Excel.Workbook)(oXL.Workbooks.Open(filepath));
                    o_LabelInfo = oWorkBook.SensitivityLabel.CreateLabelInfo();
                    System.Console.WriteLine("Setting label");
                    o_LabelInfo.AssignmentMethod = MsoAssignmentMethod.PRIVILEGED;
                    o_LabelInfo.SiteId = siteid;
                    o_LabelInfo.LabelId = labelid;
                    o_LabelInfo.LabelName = labelname;
                    oWorkBook.SensitivityLabel.SetLabel(o_LabelInfo, o_LabelInfo);
                    System.Console.WriteLine("Set label success. Saving and closing workbook.");
                    oWorkBook.Save();
                    oWorkBook.Application.Quit();
                    Marshal.ReleaseComObject(oWorkBook);
                    Marshal.ReleaseComObject(oXL);
                    result = true;

                }
                else if (Path.GetExtension(filepath).Contains(".doc"))
                {
                    System.Console.WriteLine("Word Application");
                    oW = new Word.Application { Visible = false };
                    oDocument = (Word.Document)(oW.Documents.Open(filepath));
                    o_LabelInfo = oDocument.SensitivityLabel.CreateLabelInfo();
                    System.Console.WriteLine("Setting label");
                    o_LabelInfo.AssignmentMethod = MsoAssignmentMethod.PRIVILEGED;
                    o_LabelInfo.SiteId = siteid;
                    o_LabelInfo.LabelId = labelid;
                    o_LabelInfo.LabelName = labelname;
                    oDocument.SensitivityLabel.SetLabel(o_LabelInfo, o_LabelInfo);
                    System.Console.WriteLine("Set label success. Saving and closing document.");
                    oDocument.Save();
                    oDocument.Application.Quit();
                    Marshal .ReleaseComObject(oDocument);
                    Marshal.ReleaseComObject(oW);
                    result = true;
                }
                else if (Path.GetExtension(filepath).Contains(".ppt"))
                {
                    System.Console.WriteLine("Powerpoint Application");
                    oPPT = new PowerPoint.Application();
                    oPresentation = (PowerPoint.Presentation)(oPPT.Presentations.Open(filepath));
                    o_LabelInfo = oPresentation.SensitivityLabel.CreateLabelInfo();
                    System.Console.WriteLine("Setting label");
                    o_LabelInfo.AssignmentMethod = MsoAssignmentMethod.PRIVILEGED;
                    o_LabelInfo.SiteId = siteid;
                    o_LabelInfo.LabelId = labelid;
                    o_LabelInfo.LabelName = labelname;
                    oPresentation.SensitivityLabel.SetLabel(o_LabelInfo, o_LabelInfo);
                    System.Console.WriteLine("Set label success. Saving and closing presentation.");
                    oPresentation.Save();
                    oPresentation.Application.Quit();
                    Marshal .ReleaseComObject(oPresentation);
                    Marshal.ReleaseComObject(oPPT);
                    result = true;
                }
                else
                {
                    System.Console.WriteLine("Invalid file");
                    result = false;
                    throw new Exception("File path is invalid or is not an MSO office file.");
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);

                List<dynamic> officeApplications = new List<dynamic> { oXL, oW, oPresentation };
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
                Result.Set(ctx, result);
            };
        }

        #endregion
    }
}

