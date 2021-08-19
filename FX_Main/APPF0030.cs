using System;
using System.Data;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace FX_Main
{
    public partial class APPF0030 : Office2007Form
    {
        #region ----- Enums -----


        #endregion;

        #region ----- Variables -----

        private const int mcNOCLOSEBUTTON = 0x200;

        private ISAppInterface mAppInterface;

        private System.Windows.Forms.Form mMainForm = null;

        private int mCountMenu = 0;

        private bool mIsFill = false;

        #endregion;

        #region ----- Constructor -----

        public APPF0030(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            mMainForm = pMainForm;
            mAppInterface = pAppInterface;
            isOraConnection1.OraConnectionInfo = pAppInterface.OraConnectionInfo;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void FillTreeMenu()
        {
            mIsFill = false;

            isTreeView1.HelpTextControl.ForeColor = System.Drawing.Color.White;
            isTreeView1.HelpTextControl.BackColor = System.Drawing.Color.FromArgb(0, 192, 192);

            try
            {
                idaNavigatorMenuAll.Fill();

                idaNavigatorMenuEntryAll.Fill();

                mIsFill = true;
            }
            catch
            {
            }
        }

        private void TreeMenuShow()
        {
            ISTreeView.TreeNode vMenuNode = new ISTreeView.TreeNode();
            isTreeView1.BeginUpdate();
            NavigatorMenuShow(-1, vMenuNode);
            isTreeView1.EndUpdate();
        }

        private void NavigatorMenuShow(int pMenuID, ISTreeView.TreeNode pMenuNode)
        {
            int vCountMenu = 0;

            System.Data.DataRow[] vRows1;
            if (pMenuID == -1)
            {
                vRows1 = idaNavigatorMenuAll.OraSelectData.Select("MENU_SEQ > 0");
            }
            else
            {
                string vQueryString1 = string.Format("MENU_ID = {0}", pMenuID);
                vRows1 = idaNavigatorMenuAll.OraSelectData.Select(vQueryString1);
            }

            foreach (DataRow row1 in vRows1)
            {
                if (pMenuID == -1)
                {
                    pMenuNode = new ISTreeView.TreeNode();
                    pMenuNode.Text = row1["MENU_PROMPT"].ToString();
                    isTreeView1.Nodes.Add(pMenuNode);

                    pMenuNode.Expand();
                }

                //------------------------------
                //[Tree_Node_Expand]
                //------------------------------
                if (row1["MENU_NAME"].ToString() == "ASH")
                {
                    pMenuNode.Expand();
                    vCountMenu = 0;
                }
                //------------------------------

                

                string vQueryString2 = string.Format("MENU_ID = {0}", row1["MENU_ID"]);
                System.Data.DataRow[] vRows2 = idaNavigatorMenuEntryAll.OraSelectData.Select(vQueryString2);

                foreach (DataRow row2 in vRows2)
                {
                    vCountMenu++;

                    ISTreeView.TreeNode vMenuEntryNode = new ISTreeView.TreeNode();
                    vMenuEntryNode.Text = row2["ENTRY_PROMPT"].ToString();
                    vMenuEntryNode.NodeValue = row2["ASSEMBLY_INFO_ID"];
                    vMenuEntryNode.TagObject = row2["ASSEMBLY_ID"];
                    vMenuEntryNode.Tag = row2["ASSEMBLY_FILE_NAME"];
                    vMenuEntryNode.HelpText = string.Format("{0}[{1}]", row2["ASSEMBLY_ID"].ToString(), row2["ASSEMBLY_INFO_ID"].ToString());
                    pMenuNode.Nodes.Add(vMenuEntryNode);
                    if (!row2["SUB_MENU_ID"].Equals(DBNull.Value))
                    {
                        NavigatorMenuShow(Convert.ToInt32(row2["SUB_MENU_ID"]), vMenuEntryNode); //Àç±Í
                    }
                }
            }
        }

        #endregion;

        #region ----- Node Search Methods ----

        private void ChildNodeSearch(Syncfusion.Windows.Forms.Tools.TreeNodeAdv pNode)
        {
            if (pNode.HasChildren && pNode.Nodes.Count > 0)
            {
                foreach (Syncfusion.Windows.Forms.Tools.TreeNodeAdv vSearchChildNode in pNode.Nodes)
                {
                    if (vSearchChildNode.IsSelected)
                    {
                        int vS1 = isTreeView1.VScrollPos;
                        isTreeView1.VScrollPos = mCountMenu;
                        break;
                    }
                    else
                    {
                        ChildNodeSearch(vSearchChildNode);
                    }
                }
            }
        }

        private void SearchNode()
        {
            foreach (Syncfusion.Windows.Forms.Tools.TreeNodeAdv vSearchNode in isTreeView1.Nodes)
            {
                if (vSearchNode.IsSelected)
                {
                    break;
                }
                else
                {
                    ChildNodeSearch(vSearchNode);
                }
            }
        }

        #endregion;

        #region ----- User Methods ----

        protected override CreateParams CreateParams
        {
            get
            {
                System.Windows.Forms.CreateParams vCreateParams = base.CreateParams;
                vCreateParams.ClassStyle = vCreateParams.ClassStyle | mcNOCLOSEBUTTON;

                return vCreateParams;
            }
        }

        #endregion;

        #region ----- Assembly Version Get -----

        private string GetAssemblyVersion(string pAssemblyFineName)
        {
            string vVersionAssembly = string.Empty;
            System.Diagnostics.FileVersionInfo vFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(pAssemblyFineName);
            vVersionAssembly = vFileVersionInfo.FileVersion;

            return vVersionAssembly;
        }

        #endregion;

        #region ----- FormLoading Methods -----

        private void ProcessStart(string pPath, ISTreeView.TreeNode pSelectedNode)
        {
            try
            {
                object vObject = null;
                string vAssemblyVersion = string.Empty;
                string vAssemblyFileName = string.Empty;
                string vAssemblyId = string.Empty;
                string vEntryPrompt = string.Empty;

                vEntryPrompt = pSelectedNode.Text;

                vObject = pSelectedNode.TagObject;
                if (vObject == null)
                {
                    MessageBoxAdv.Show("Assembly File Not Found");
                    return;
                }
                vAssemblyId = vObject as string;

                vObject = pSelectedNode.Tag;
                if (vObject == null)
                {
                    MessageBoxAdv.Show("Assembly File Not Found");
                    return;
                }
                vAssemblyFileName = vObject as string;

                bool isNull = string.IsNullOrEmpty(pPath);
                if (isNull == true)
                {
                    string vPath = @"..\..\..";
                    vAssemblyFileName = string.Format("{0}\\{1}\\bin\\Debug\\{2}", vPath, vAssemblyId, vAssemblyFileName);
                }
                else
                {
                    vAssemblyFileName = string.Format("{0}{1}", pPath, vAssemblyFileName);
                }

                System.IO.FileInfo vFileInfo = new System.IO.FileInfo(vAssemblyFileName);
                if (vFileInfo.Exists == true)
                {
                    vAssemblyVersion = GetAssemblyVersion(vAssemblyFileName);
                    System.Reflection.Assembly vAssembly = System.Reflection.Assembly.LoadFrom(vAssemblyFileName);

                    string vNameSpace = vAssembly.GetName().Name;
                    string vNameSpaceClass = string.Format("{0}.{1}", vNameSpace, vNameSpace);
                    Type vType = vAssembly.GetType(vNameSpaceClass);

                    if (vType != null)
                    {
                        object[] vParam = new object[2];
                        vParam[0] = this.MdiParent;
                        vParam[1] = isAppInterfaceAdv1.AppInterface;

                        object vCreateInstance = Activator.CreateInstance(vType, vParam);
                        Form vForm = vCreateInstance as Form;
                        string vCaption = string.Format("{0}[{1}] - {2}", vEntryPrompt, vAssemblyId, vAssemblyVersion);
                        vForm.Text = vCaption;
                        vForm.Show();
                    }
                    else
                    {
                        MessageBoxAdv.Show("Form Namespace Error");
                    }
                }
                else
                {
                    MessageBoxAdv.Show("Assembly File Not Found");
                }
            }
            catch (Exception ex)
            {
                MessageBoxAdv.Show(ex.Message);
            }
        }

        //private void ProcessStart(string pPath, string pAssemblyName)
        //{
        //    try
        //    {
        //        string vAssemblyVersion = string.Empty;
        //        string vAssemblyFileName = string.Empty;
        //        bool isNull = string.IsNullOrEmpty(pPath);
        //        if (isNull == true)
        //        {
        //            string vPath = @"..\..\..";
        //            int viCutStart = pAssemblyName.LastIndexOf(".");
        //            string vProjectName = pAssemblyName.Substring(0, viCutStart);
        //            vAssemblyFileName = string.Format("{0}\\{1}\\bin\\Debug\\{2}", vPath, vProjectName, pAssemblyName);
        //        }
        //        else
        //        {
        //            vAssemblyFileName = string.Format("{0}{1}", pPath, pAssemblyName);
        //        }
                
        //        System.IO.FileInfo vFileInfo = new System.IO.FileInfo(vAssemblyFileName);
        //        if (vFileInfo.Exists == true)
        //        {
        //            vAssemblyVersion = GetAssemblyVersion(vAssemblyFileName);
        //            System.Reflection.Assembly vAssembly = System.Reflection.Assembly.LoadFrom(vAssemblyFileName);

        //            string vNameSpace = vAssembly.GetName().Name;
        //            string vNameSpaceClass = string.Format("{0}.{1}", vNameSpace, vNameSpace);
        //            Type vType = vAssembly.GetType(vNameSpaceClass);

        //            if (vType != null)
        //            {
        //                object[] vParam = new object[2];
        //                vParam[0] = this.MdiParent;
        //                vParam[1] = isAppInterfaceAdv1.AppInterface;

        //                object vCreateInstance = Activator.CreateInstance(vType, vParam);
        //                Form vForm = vCreateInstance as Form;
        //                string vCaption = string.Format("{0} - {1}", vForm.Text, vAssemblyVersion);
        //                vForm.Text = vCaption;
        //                vForm.Show();
        //            }
        //            else
        //            {
        //                MessageBoxAdv.Show("Form Namespace Error");
        //            }
        //        }
        //        else
        //        {
        //            MessageBoxAdv.Show("Assembly File Not Found");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBoxAdv.Show(ex.Message);
        //    }
        //}

        #endregion;

        #region ----- Events -----

        private void APPF0030_Load(object sender, EventArgs e)
        {
            FillTreeMenu();
        }

        private void APPF0030_Shown(object sender, EventArgs e)
        {
            if (mIsFill == true)
            {
                TreeMenuShow();
                SearchNode();
            }
        }

        private void buttonAdv1_Click(object sender, EventArgs e)
        {
            isTreeView1.Nodes.Clear();
            FillTreeMenu();
            TreeMenuShow();

            SearchNode();
        }

        private void isTreeView1_DoubleClick(object sender, EventArgs e)
        {
            if (isTreeView1.SelectedNode != null)
            {
                ISTreeView.TreeNode vSelectedNode = (ISTreeView.TreeNode)isTreeView1.SelectedNode;

                if (vSelectedNode.Value != null && vSelectedNode.Value != DBNull.Value)
                {
                    this.Cursor = Cursors.WaitCursor;
                    Application.DoEvents();
                    isDataCommand1.SetCommandParamValue("W_ASSEMBLY_INFO_ID", vSelectedNode.NodeValue);
                    isDataCommand1.ExecuteNonQuery();

                    string vPath = string.Empty;
                    object vObject2 = isDataCommand1.GetCommandParamValue("O_ASSEMBLY_PATH");
                    bool isConvert2 = vObject2 is string;
                    if (isConvert2 == true)
                    {
                        vPath = vObject2 as string;
                    }

                    object vObject1 = isDataCommand1.GetCommandParamValue("O_ASSEMBLY_FILE_NAME");
                    
                    bool isConvert1 = vObject1 is string;
                    if (isConvert1 == true)
                    {
                        string vAssemblyFileName = vObject1 as string;
                        bool isNull = string.IsNullOrEmpty(vAssemblyFileName);
                        if (isNull == false)
                        {
                            //ProcessStart(vPath, vAssemblyFileName);
                            ProcessStart(vPath, vSelectedNode);
                        }
                    }
                    else
                    {
                        MessageBoxAdv.Show("This program is not available.");
                    }

                    this.Cursor = Cursors.Default;
                    Application.DoEvents();
                }
            }
        }

        #endregion;
    }
}