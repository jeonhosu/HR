using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;


namespace EAPF0108
{
    public partial class EAPF0108 : Office2007Form
    {
        #region ----- Variables -----

        private object mRadioValue_1_Search = 'N';
        private object mRadioValue_2_Insert = 'N';

        #endregion;

        #region ----- Constructor -----

        public EAPF0108()
        {
            InitializeComponent();
        }

        public EAPF0108(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- User Methods -----

        private DateTime GetDateTime()
        {
            DateTime vDateTime = DateTime.Today;

            try
            {
                IDC_GET_DATE.ExecuteNonQuery();
                object vObject = IDC_GET_DATE.GetCommandParamValue("X_LOCAL_DATE");

                bool isConvert = vObject is DateTime;
                if (isConvert == true)
                {
                    vDateTime = (DateTime)vObject;
                }
            }
            catch (Exception ex)
            {
                string vMessage = ex.Message;
                vDateTime = new DateTime(9999, 12, 31, 23, 59, 59);
            }
            return vDateTime;
        }

        #endregion;

        #region ----- Private Methods ----

        private bool IsNull(object pObject)
        {
            bool isNull = false;
            bool isConvert = pObject is string;
            if (isConvert == true)
            {
                string vString = pObject as string;
                isNull = string.IsNullOrEmpty(vString.Trim());
            }
            return isNull;
        }

        private void DefaultSetCheckBox2()
        {
            ENABLED_FLAG.CheckedState = ISUtil.Enum.CheckedState.Checked;
        }

        private void DefaultSetDateTime1()
        {
            DateTime vGetDateTime = GetDateTime();

            //ISCommonUtil.ISFunction.ISDateTime vDate = new ISCommonUtil.ISFunction.ISDateTime();
            //DateTime vMonthFirstDay = vDate.ISMonth_1st(vGetDateTime);

            EFFECTIVE_DATE_FR.EditValue = vGetDateTime;
        }

        #region ----- Is User Select -----
        //사용자 추가시
        //사용자 형식의 Super User, Outside User에서는
        //HRM_PERSON_MASTER.PERSON_ID 가 필요없고,
        //General User 에서는
        //꼭 HRM_PERSON_MASTER.PERSON_ID 가 필수 이다.
        private bool IsUserSelect(ISEditAdv pEditAdv1, ISEditAdv pEditAdv2)
        {
            //mRadioValue_2_Insert
            bool isAble = false;
            string vUserType = mRadioValue_2_Insert as string;

            if (vUserType == "S" || vUserType == "B") //Super User
            {
                isAble = true;
            }
            else if (vUserType == "A") //General User
            {
                if (pEditAdv2.EditValue != null)
                {
                    bool isConvert1 = pEditAdv1.EditValue is string;
                    bool isConvert2 = pEditAdv2.EditValue is decimal;
                    if (isConvert1 == true && isConvert2 == true)
                    {
                        string vText1 = pEditAdv1.EditValue as string;
                        decimal vValue = (decimal)pEditAdv2.EditValue;
                        bool isNull1 = string.IsNullOrEmpty(vText1);
                        if (isNull1 != true && vValue != 0)
                        {
                            isAble = true;
                        }
                    }
                }
            }

            return isAble;
        }

        #endregion;

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    int vIndexRow = IGR_EAPP_USER.RowIndex;
                    int vIndexColumn = IGR_EAPP_USER.ColIndex;

                    SearchFromDataAdapter();

                    int vCountRow = IGR_EAPP_USER.RowCount;
                    if (vCountRow == 1)
                    {
                        vIndexRow = 1;
                    }

                    IGR_EAPP_USER.CurrentCellMoveTo(vIndexRow, vIndexColumn);
                    IGR_EAPP_USER.Focus();
                    IGR_EAPP_USER.CurrentCellActivate(vIndexRow, vIndexColumn);

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_EAPP_USER.IsFocused == true)
                    {
                        IDA_EAPP_USER.AddOver();

                        DefaultSetCheckBox2();
                        DefaultSetDateTime1();

                        RB_GENERAL_USER.Checked = true;
                    }
                    else if (IDA_USER_AUTHORITY_GROUP.IsFocused == true)
                    {
                        IDA_USER_AUTHORITY_GROUP.AddOver();
                        IGR_USER_AUTHORITY_GROUP.SetCellValue("ENABLED_FLAG", "Y");
                    }
                    else if (IDA_USER_PRG_AUTHORITY_ADD.IsFocused == true)
                    {
                        IDA_USER_PRG_AUTHORITY_ADD.AddOver();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_EAPP_USER.IsFocused == true)
                    {
                        IDA_EAPP_USER.AddUnder();

                        DefaultSetCheckBox2();
                        DefaultSetDateTime1();

                        RB_GENERAL_USER.Checked = true;
                    }
                    else if (IDA_USER_AUTHORITY_GROUP.IsFocused == true)
                    {
                        IDA_USER_AUTHORITY_GROUP.AddUnder();
                        IGR_USER_AUTHORITY_GROUP.SetCellValue("ENABLED_FLAG", "Y");
                    }
                    else if (IDA_USER_PRG_AUTHORITY_ADD.IsFocused == true)
                    {
                        IDA_USER_PRG_AUTHORITY_ADD.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_EAPP_USER.IsFocused == true)
                    {
                        //bool IsAble = IsUserSelect(W_USER, W_DEPT_MASTER);
                        //if (IsAble == false)
                        //{
                        //    string vMessageGet = isMessageAdapter1.ReturnText("EAPP_10037"); //사용자를 선택 하세요!
                        //    string vMessageString = string.Format("{0}", vMessageGet);
                        //    MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //    return;
                        //}

                        object vAuthorityType = vAuthorityType = "S";

                        IDA_EAPP_USER.SetInsertParamValue("P_USER_TYPE", mRadioValue_2_Insert);
                        IDA_EAPP_USER.SetInsertParamValue("P_AUTHORITY_TYPE", vAuthorityType);

                        //-----------------------------------------------------------

                        IDA_EAPP_USER.SetUpdateParamValue("P_USER_TYPE", mRadioValue_2_Insert);
                        IDA_EAPP_USER.SetUpdateParamValue("P_AUTHORITY_TYPE", vAuthorityType);
                        IDA_EAPP_USER.Update();
                    }
                    else if (IDA_USER_AUTHORITY_GROUP.IsFocused == true)
                    {
                        IDA_USER_AUTHORITY_GROUP.Update();
                        IDA_USER_AUTHORITY_GROUP.Fill();
                    }
                    else if (IDA_USER_PRG_AUTHORITY_ADD.IsFocused == true)
                    {
                        IDA_USER_PRG_AUTHORITY_ADD.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_EAPP_USER.IsFocused)
                    {
                        IDA_EAPP_USER.Cancel();
                    }
                    else if (IDA_USER_AUTHORITY_GROUP.IsFocused == true)
                    {
                        IDA_USER_AUTHORITY_GROUP.Cancel();
                    }
                    else if (IDA_USER_PRG_AUTHORITY_ADD.IsFocused == true)
                    {
                        IDA_USER_PRG_AUTHORITY_ADD.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (IDA_EAPP_USER.IsFocused)
                    {
                        IDA_EAPP_USER.Delete();
                    }
                    else if (IDA_USER_AUTHORITY_GROUP.IsFocused == true)
                    {
                        IDA_USER_AUTHORITY_GROUP.Delete();
                    }
                    else if (IDA_USER_PRG_AUTHORITY_ADD.IsFocused == true)
                    {
                        IDA_USER_PRG_AUTHORITY_ADD.Delete();
                    }
                }
            }
        }

        #endregion;

        private void SearchFromDataAdapter()
        {
            bool isNull = false;
            object vOjbect = false;

            vOjbect = W_USER.EditValue;
            isNull = IsNull(vOjbect);
            if (isNull != true)
            {
                IDA_EAPP_USER.SetSelectParamValue("W_DESCRIPTION", vOjbect);
            }
            else
            {
                IDA_EAPP_USER.SetSelectParamValue("W_DESCRIPTION", null);
            }

            vOjbect = W_DEPT_MASTER.EditValue;
            isNull = IsNull(W_DEPT_MASTER.EditValue);
            if (isNull != true)
            {
                IDA_EAPP_USER.SetSelectParamValue("W_DEPT_NAME", vOjbect);
            }
            else
            {
                IDA_EAPP_USER.SetSelectParamValue("W_DEPT_NAME", null);
            }

            isNull = IsNull(mRadioValue_1_Search);
            if (isNull != true)
            {
                IDA_EAPP_USER.SetSelectParamValue("W_USER_TYPE", mRadioValue_1_Search);
            }
            else
            {
                IDA_EAPP_USER.SetSelectParamValue("W_USER_TYPE", null);
            }

            IDA_EAPP_USER.Fill();
        }

        private void W_RB_GENERAL_USER_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;

            if (vRadio.Checked == true)
            {
                mRadioValue_1_Search = vRadio.RadioCheckedString;
            }
        }

        private void RB_GENERAL_USER_CheckChanged(object sender, EventArgs e)
        {
            ISRadioButtonAdv vRadio = sender as ISRadioButtonAdv;

            if (vRadio.Checked == true)
            {
                USER_TYPE.EditValue = vRadio.RadioCheckedString;
                mRadioValue_2_Insert = vRadio.RadioCheckedString;
            }
        }

        private void EAPF0108_Load(object sender, EventArgs e)
        {
            IDA_EAPP_USER.FillSchema();
    
            mRadioValue_1_Search = 'N';
            mRadioValue_2_Insert = 'A';
            
        }

        private void EAPF0108_Shown(object sender, EventArgs e)
        {
            W_USER.Focus();
        }

        private void IDA_EAPP_USER_NewRowMoved(object pSender, ISBindingEventArgs pBindingManager)
        {

        }

    }
}