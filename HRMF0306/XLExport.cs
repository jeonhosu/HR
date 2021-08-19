using System;
using System.Windows.Forms;

namespace HRMF0336
{
    public partial class XLExport : Form
    {
        #region ----- Variables -----

        private string mMessageError = string.Empty;

        private XL.XLPrint mExport = null;

        #endregion;

        #region ----- Property -----


        #endregion;

        #region ----- Constructor -----

        public XLExport()
        {
            InitializeComponent();

            mExport = new XL.XLPrint();
        }

        #endregion;

        #region ----- Convert Date Methods ----

        private object ConvertDate(object pObject)
        {
            object vObject = null;

            try
            {
                if (pObject != null)
                {
                    bool IsConvert = pObject is System.DateTime;
                    if (IsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        string vTextDateTimeLong = vDateTime.ToString("yyyy-MM-dd", null);
                        string vTextDateTimeShort = vDateTime.ToShortDateString();
                        vObject = vTextDateTimeLong;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return vObject;
        }

        #endregion;

        #region ----- XLDispose -----

        private void XLDispose()
        {
            mExport.XLOpenFileClose();
            mExport.XLClose();
        }

        #endregion;

        #region ----- XL File Open -----

        private bool XLFileOpen(string pXLOpenFileName)
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mExport.XLOpenFile(pXLOpenFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);
            }

            return IsOpen;
        }

        #endregion;

        #region ----- Save Methods ----

        private void Save(string pSaveFileName)
        {
            try
            {
                mExport.XLSave(pSaveFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);

                XLDispose();
            }
        }

        #endregion;

        #region ----- Sheet Name Methods ----

        private void SetSheetName(int pIndexSheet, string pSheetName)
        {
            string vSheetName = string.Empty;

            //Excel_Sheet_Name_Not_Char : /, \, *, [, ], :, ?
            try
            {
                vSheetName = pSheetName;
                vSheetName = vSheetName.Replace("/", "");
                vSheetName = vSheetName.Replace("\\", "");
                vSheetName = vSheetName.Replace("*", "");
                vSheetName = vSheetName.Replace("[", "");
                vSheetName = vSheetName.Replace("]", "");
                vSheetName = vSheetName.Replace(":", "");
                vSheetName = vSheetName.Replace("?", "");

                vSheetName = string.Format("{0}", vSheetName);
                mExport.XLSheetName(pIndexSheet, vSheetName);
            }
            catch
            {
            }
        }

        #endregion;

        #region ----- Header Columns Methods ----

        private void XLHeaderColumns(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pTerritory)
        {
            int vXLine = 2;
            int vXColumn = 1;
            int vCountColumn = pGrid.ColCount;

            object vObject = null;

            try
            {
                if (vCountColumn < 1)
                {
                    return;
                }

                mExport.XLActiveSheet(1);

                vXLine = 4;

                //Header Columns
                for (int vCol = 0; vCol < vCountColumn; vCol++)
                {
                    vObject = pGrid.GridAdvExColElement[vCol].Visible;
                    int vVisible = (int)vObject;

                    // 폼 타이틀 출력되는 부분
                    if (vVisible == 1)
                    {
                        switch (pTerritory)
                        {
                            case 1: //Default
                                vObject = pGrid.GridAdvExColElement[vCol].HeaderElement[0].Default;
                                mExport.XLSetCell(vXLine, vXColumn, vObject);
                                break;
                            case 2: //KR
                                vObject = pGrid.GridAdvExColElement[vCol].HeaderElement[0].TL1_KR;
                                mExport.XLSetCell(vXLine, vXColumn, vObject);
                                break;
                        }

                        vXColumn++;
                    }
                }
                // 사용자 이름 출력되는 부분
                /*
                System.Security.Principal.WindowsIdentity vUser;
                vUser = System.Security.Principal.WindowsIdentity.GetCurrent();
                string vUserInfo = string.Format("{0}", vUser.Name);
                mExport.XLSetCell(1, 1, vUserInfo);
                */
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);

                XLDispose();
            }
        }

        #endregion;

        #region ----- Excel Wirte Methods ----

        private bool XLWirte(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {
            bool isWrite = false;
            object vObject = null;
            int vXLine = 5;
            int vXColumn = 1;

            try
            {
                mExport.XLActiveSheet(1);

                int vTotalRow = pGrid.RowCount;
                int vCountColumn = pGrid.ColCount;

                for (int vRow = 0; vRow < vTotalRow; vRow++)
                {
                    //pGrid.CurrentCellMoveTo(vRow, 0);
                    //pGrid.Focus();
                    //pGrid.CurrentCellActivate(vRow, 0);

                    for (int vCol = 0; vCol < vCountColumn; vCol++)
                    {
                        vObject = pGrid.GridAdvExColElement[vCol].Visible;
                        int vVisible = (int)vObject;

                        if (vVisible == 1)
                        {
                            vObject = pGrid.GetCellValue(vRow, vCol);
                            mExport.XLSetCell(vXLine, vXColumn, vObject);

                            vXColumn++;
                        }
                    }

                    isProgressBar1.BarFillPercent = (Convert.ToSingle(vXLine) / Convert.ToSingle(vTotalRow)) * 100F;
                    vXLine++;
                    vXColumn = 1;
                }

                mExport.XLColumnAutoFit(1, 1, vTotalRow, vCountColumn);

                isWrite = true;
            }
            catch (System.Exception ex)
            {
                mMessageError = string.Format("{0}", ex.Message);

                XLDispose();
            }

            return isWrite;
        }

        #endregion;

        #region ----- Excel Export Methods ----

        public bool ExcelExport(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pTerritory, string pOpenExcelFileName, string pSaveExcelFileName, string pTitle, System.Windows.Forms.Form pForm)
        {
            bool isWrite = false;

            mExport = new XL.XLPrint();

            bool isOpen = XLFileOpen(pOpenExcelFileName);

            if (isOpen == true)
            {
                XLHeaderColumns(pGrid, pTerritory);

                this.Show(pForm);
                this.ClientSize = new System.Drawing.Size(370, 32);
                this.BringToFront();
                System.Windows.Forms.Application.DoEvents();

                isWrite = XLWirte(pGrid);

                //Title
                //mExport.XLSetCell(2, 1, pTitle); //행, 열, 값

                SetSheetName(1, pTitle);

                if (isWrite == true)
                {
                    Save(pSaveExcelFileName);
                }

                this.Hide();
                this.SendToBack();
            }

            XLDispose();

            return isWrite;
        }

        #endregion;
    }
}
