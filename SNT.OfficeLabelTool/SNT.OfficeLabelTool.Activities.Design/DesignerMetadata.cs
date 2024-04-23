using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using SNT.OfficeLabelTool.Activities.Design.Designers;
using SNT.OfficeLabelTool.Activities.Design.Properties;

namespace SNT.OfficeLabelTool.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(GetSensitivityLabel), categoryAttribute);
            builder.AddCustomAttributes(typeof(GetSensitivityLabel), new DesignerAttribute(typeof(GetSensitivityLabelDesigner)));
            builder.AddCustomAttributes(typeof(GetSensitivityLabel), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(SetSensitivityLabel), categoryAttribute);
            builder.AddCustomAttributes(typeof(SetSensitivityLabel), new DesignerAttribute(typeof(SetSensitivityLabelDesigner)));
            builder.AddCustomAttributes(typeof(SetSensitivityLabel), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
