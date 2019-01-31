using System.AddIn;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;

////////////////////////////////////////////////////////////////////////////////
//
// File: WorkspaceAddIn.cs
//
// Comments:
//
// Notes: 
//
// Pre-Conditions: 
//
////////////////////////////////////////////////////////////////////////////////
namespace Labor_Automation
{
    public class WorkspaceAddIn : Panel, IWorkspaceComponent2
    {
        private Label label1;

        /// <summary>
        /// The current workspace record context.
        /// </summary>
        private IRecordContext _recordContext;
        public ICustomObject _laborRecord { get; set; }
        public IOrganization _orgRecord { get; set; }
        /// <summary>
        /// Default constructor.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        public WorkspaceAddIn(bool inDesignMode, IRecordContext RecordContext)
        {
            if (!inDesignMode)
            {
                _recordContext = RecordContext;
            }
            else
            {
                InitializeComponent();
            }
        }

        /// <summary>
        /// Method called by the Add-In framework to initialize it in design mode.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Margin = new System.Windows.Forms.Padding(10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Decimal Conversion upto 2 digit";
            // 
            // WorkspaceAddIn
            // 
            this.Controls.Add(this.label1);
            this.Size = new System.Drawing.Size(20, 10);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #region IAddInControl Members

        /// <summary>
        /// Method called by the Add-In framework to retrieve the control.
        /// </summary>
        /// <returns>The control, typically 'this'.</returns>
        public Control GetControl()
        {
            return this;
        }

        #endregion

        #region IWorkspaceComponent2 Members

        /// <summary>
        /// Sets the ReadOnly property of this control.
        /// </summary>
        public bool ReadOnly { get; set; }

        /// <summary>
        /// Method which is called when any Workspace Rule Action is invoked.
        /// </summary>
        /// <param name="ActionName">The name of the Workspace Rule Action that was invoked.</param>
        public void RuleActionInvoked(string ActionName)
        {
            _laborRecord = _recordContext.GetWorkspaceRecord(_recordContext.WorkspaceTypeName) as ICustomObject;
            if (ActionName.Contains("DecimalConvert"))
            {
                string fielName = ActionName.Split('@')[1];//Get the field name, which was changed
                string fieldVal = GetFieldValue(fielName);//Get value
                string convertedFieldValue = Decimalconversion(fieldVal, fielName);//Convert it 2 decimal
                if (convertedFieldValue != "")
                    SetFieldValue(fielName, convertedFieldValue);//set the value
                _recordContext.RefreshWorkspace();//refresh the workspace
            }
            if(ActionName.Contains("getOrgRate"))
            {
                string orgLaborRate = getOrgField("CO", "labor_rate");
                if (orgLaborRate == "")
                    MessageBox.Show(_orgRecord.Name+" doesn't have any labor rate");
                if (ActionName == "getOrgRateClaim")
                {
                    SetFieldValue("labor_rate_adj", orgLaborRate);
                } else
                {
                    SetFieldValue("labor_rate_rqstd_org", orgLaborRate);
                }
            }

            if (ActionName.Contains("@CopyTO@"))
            {
                string[] fields = ActionName.Split(new string[] { "@CopyTO@" }, StringSplitOptions.None);
                string sourceFielName = fields[0]; //Get the field name, which was changed
                string destinationFieldName = fields[1];
                string fieldVal = GetFieldValue(sourceFielName);//Get value
                SetFieldValue(destinationFieldName, fieldVal);//set the value
                _recordContext.RefreshWorkspace();//refresh the workspace
            }
        }

        /// <summary>
        /// Method which is called when any Workspace Rule Condition is invoked.
        /// </summary>
        /// <param name="ConditionName">The name of the Workspace Rule Condition that was invoked.</param>
        /// <returns>The result of the condition.</returns>
        public string RuleConditionInvoked(string ConditionName)
        {
            return string.Empty;
        }
        /// <summary>
        /// Method which is called to get value of a field.
        /// </summary>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public string GetFieldValue(string fieldName)
        {
            IList<IGenericField> fields = _laborRecord.GenericFields;
            if (null != fields)
            {
                foreach (IGenericField field in fields)
                {
                    if (field.Name.Equals(fieldName))
                    {
                        if (field.DataValue.Value != null)
                            return field.DataValue.Value.ToString();
                    }
                }
            }
            return "";
        }
        /// <summary>
        /// Method which is use to set value to a field using record Context 
        /// </summary>
        /// <param name="fieldName">field name</param>
        /// <param name="value">value of field</param>
        public void SetFieldValue(string fieldName, string value)
        {
            IList<IGenericField> fields = _laborRecord.GenericFields;
            if (null != fields)
            {
                foreach (IGenericField field in fields)
                {
                    if (field.Name.Equals(fieldName))
                    {
                        switch (field.DataType)
                        {
                            case RightNow.AddIns.Common.DataTypeEnum.STRING:
                                field.DataValue.Value = value;
                                break;
                        }
                    }
                }
            }
            return;
        }
        /// <summary>
        /// Method which is called to get value of a custom field of Org record.
        /// </summary>
        /// <param name="packageName">The name of the package.</param>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public string getOrgField(string packageName, string fieldName)
        {
            string value = "";
            _orgRecord = (IOrganization)_recordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Organization);
            if (_orgRecord == null)
                return "";

            IList<ICustomAttribute> orgCustomAttributes = _orgRecord.CustomAttributes;

            foreach (ICustomAttribute val in orgCustomAttributes)
            {
                if (val.PackageName == packageName)//if package name matches
                {
                    if (val.GenericField.Name == packageName + "$" + fieldName)//if field matches
                    {
                        if (val.GenericField.DataValue.Value != null)
                        {
                            value = val.GenericField.DataValue.Value.ToString();
                            break;
                        }
                    }
                }
            }
            return value;
        }
        /// <summary>
        /// Function to convert string value to decimal upto 2 digit
        /// </summary>
        /// <param name="val"></param>
        /// <param name="fieldName"></param>
        /// <returns>Decimal Value upto 2 digit</returns>
        public string Decimalconversion(string val, string fieldName)
        {
            Decimal output = 0;
            try
            {
                if (Regex.Matches(val, @"[a-zA-Z]").Count > 0)
                {
                    MessageBox.Show("Value in field " + fieldName + " " + "should be Numeric", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else
                {
                    if (!(val.Contains(".")))
                    {
                        val = val + ".00";
                    }
                    if ((val.Contains(".")))
                    {
                        int length = val.Substring(val.IndexOf(".")).Length;
                        if (length == 2)
                        {
                            val = val + "0";
                        }
                    }
                    output = Math.Round(Convert.ToDecimal(val), 2);
                    return output.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return "";
        }
        #endregion

    }

    [AddIn("Workspace Factory AddIn", Version = "1.0.0.0")]
    public class WorkspaceAddInFactory : IWorkspaceComponentFactory2
    {
        #region IWorkspaceComponentFactory2 Members

        /// <summary>
        /// Method which is invoked by the AddIn framework when the control is created.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        /// <returns>The control which implements the IWorkspaceComponent2 interface.</returns>
        public IWorkspaceComponent2 CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceAddIn(inDesignMode, RecordContext);
        }

        #endregion

        #region IFactoryBase Members

        /// <summary>
        /// The 16x16 pixel icon to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public Image Image16
        {
            get { return Properties.Resources.AddIn16; }
        }

        /// <summary>
        /// The text to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Text
        {
            get { return "Labor Automation"; }
        }

        /// <summary>
        /// The tooltip displayed when hovering over the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Tooltip
        {
            get { return "This Add-in converts Price fields to Decimal values"; }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            return true;
        }

        #endregion
    }
}