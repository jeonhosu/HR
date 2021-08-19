using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace FX_Main
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            string[] vSplit = new string[25];

            //---------------------------------------------------------------------
            //// Prod flexcom.
            //Prod
            vSplit[0] = "218.156.85.220";                   //Oracle_Host
            vSplit[1] = "1521";                           //Oracle_Port
            vSplit[2] = "SEILPROD";                          //Oracle_ServiceName
            vSplit[3] = "APPS";                           //Oracle_UserId
            vSplit[4] = "infoflex";                       //Oracle_Password          
            vSplit[11] = "4"; //LoginId;
            vSplit[19] = "100937"; //LoginNo                        
            vSplit[20] = "1148"; //PersonID
            vSplit[21] = "100937"; //PersonNumber

            // Test
            vSplit[0] = "192.168.10.245";                   //Oracle_Host
            vSplit[1] = "1521";                           //Oracle_Port           
            vSplit[2] = "BSKPROD ";                          //Oracle_ServiceName 
            vSplit[3] = "APPS";                           //Oracle_UserId
            vSplit[4] = "infoflex";                       //Oracle_Password          
            vSplit[11] = "4"; //LoginId;
            vSplit[19] = "12070901"; //LoginNo                        
            vSplit[20] = "3254"; //PersonID
            vSplit[21] = "12070901"; //PersonNumber
            vSplit[9] = "70";   //SOB_ID
            vSplit[10] = "701"; //ORG_ID

            ////////// Test
            //vSplit[0] = "192.168.10.10";  //"192.168.1.7"; //"1.241.249.174";                   //Oracle_Host
            //vSplit[1] = "1521";                           //Oracle_Port           
            //vSplit[2] = "MESORA";  //"FEKPROD";                          //Oracle_ServiceName 
            //vSplit[3] = "apps";                           //Oracle_UserId
            //vSplit[4] = "erp0201";                       //Oracle_Password          
            //vSplit[11] = "4"; //LoginId;
            //vSplit[19] = "100937"; //LoginNo                        
            //vSplit[20] = "191"; //PersonID
            //vSplit[21] = "100937"; //PersonNumber
            //vSplit[9] = "70";   //SOB_ID
            //vSplit[10] = "701"; //ORG_ID

            //////////// Test
            vSplit[0] = "106.251.238.99";                   //Oracle_Host
            vSplit[1] = "1521";                           //Oracle_Port           
            vSplit[2] = "NFMPROD2";                          //Oracle_ServiceName 
            vSplit[3] = "APPS";                           //Oracle_UserId
            vSplit[4] = "infoflex!@";                       //Oracle_Password          
            vSplit[11] = "3"; //LoginId;
            vSplit[19] = "100937"; //LoginNo                        
            vSplit[20] = "21933"; //이성의 // "20125"; //김종구  // "22783"; //김유진   // PersonID
            vSplit[21] = "100937"; //PersonNumber
            vSplit[9] = "80";   //SOB_ID
            vSplit[10] = "801"; //ORG_ID

            ////// Test
            //vSplit[0] = "192.168.1.7";                   //Oracle_Host
            //vSplit[1] = "1521";                           //Oracle_Port           
            //vSplit[2] = "SIKPROD";                          //Oracle_ServiceName 
            //vSplit[3] = "APPS";                           //Oracle_UserId
            //vSplit[4] = "infoflex";                       //Oracle_Password     
            //vSplit[11] = "4"; //LoginId;
            //vSplit[19] = "100937"; //LoginNo                        
            //vSplit[20] = "192"; //PersonID
            //vSplit[21] = "100937"; //PersonNumber
            //vSplit[9] = "70";   //SOB_ID
            //vSplit[10] = "701"; //ORG_ID

            //// Test
            //vSplit[0] = "14.241.92.241";                   //Oracle_Host
            //vSplit[1] = "1521";                           //Oracle_Port           
            //vSplit[2] = "NFVPROD";                          //Oracle_ServiceName 
            //vSplit[3] = "APPS";                           //Oracle_UserId
            //vSplit[4] = "infoflex!";                       //Oracle_Password     
            //vSplit[11] = "4"; //LoginId;
            //vSplit[19] = "100937"; //LoginNo                        
            //vSplit[20] = "22783"; //PersonID
            //vSplit[21] = "100937"; //PersonNumber
            //vSplit[9] = "70";   //SOB_ID
            //vSplit[10] = "701"; //ORG_ID

            //Test
            vSplit[0] = "58.151.251.160";                   //Oracle_Host
            vSplit[1] = "1521";                           //Oracle_Port           
            vSplit[2] = "NFKPROD";                          //Oracle_ServiceName 
            vSplit[3] = "APPS";                           //Oracle_UserId
            vSplit[4] = "infoflex!";                       //Oracle_Password     
            vSplit[11] = "3"; //LoginId;
            vSplit[19] = "100937"; //LoginNo                        
            vSplit[20] = "22783";  //최다솔.//"20635"; //최진 // "22855"; //서정필  // "20635"; //최진 //PersonID
            vSplit[21] = "100937"; //PersonNumber
            vSplit[9] = "80";   //SOB_ID
            vSplit[10] = "801"; //ORG_ID

            //// Test
            //vSplit[0] = "14.241.92.241";                   //Oracle_Host
            //vSplit[1] = "1521";                           //Oracle_Port           
            //vSplit[2] = "NFVPROD";                          //Oracle_ServiceName 
            //vSplit[3] = "APPS";                           //Oracle_UserId
            //vSplit[4] = "infoflex!";                       //Oracle_Password     
            //vSplit[11] = "4"; //LoginId;
            //vSplit[19] = "100937"; //LoginNo                        
            //vSplit[20] = "192"; //PersonID
            //vSplit[21] = "100937"; //PersonNumber
            //vSplit[9] = "70";   //SOB_ID
            //vSplit[10] = "701"; //ORG_ID


            //// Test
            //vSplit[0] = "14.238.9.67";                   //Oracle_Host
            //vSplit[1] = "1521";                           //Oracle_Port           
            //vSplit[2] = "NFVPROD";                          //Oracle_ServiceName 
            //vSplit[3] = "APPS";                           //Oracle_UserId
            //vSplit[4] = "infoflex!";                       //Oracle_Password     
            //vSplit[11] = "3"; //LoginId;
            //vSplit[19] = "100937"; //LoginNo                        
            //vSplit[20] = "-1"; //PersonID
            //vSplit[21] = "100937"; //PersonNumber
            //vSplit[9] = "70";   //SOB_ID
            //vSplit[10] = "701"; //ORG_ID


            ////// Test
            //vSplit[0] = "172.16.160.15";  //"192.168.1.7"; //"1.241.249.174";                   //Oracle_Host
            //vSplit[1] = "1521";                           //Oracle_Port           
            //vSplit[2] = "SIVPROD";  //"FEKPROD";                          //Oracle_ServiceName 
            //vSplit[3] = "apps";                           //Oracle_UserId
            //vSplit[4] = "infoflex";                       //Oracle_Password          
            //vSplit[11] = "4"; //LoginId;
            //vSplit[19] = "100937"; //LoginNo                        
            //vSplit[20] = "12322"; //PersonID
            //vSplit[21] = "100937"; //PersonNumber
            //vSplit[9] = "90";   //SOB_ID
            //vSplit[10] = "901"; //ORG_ID

            vSplit[5] = "106.251.238.99";                 //FTP_Host;
            vSplit[6] = "1502";                             //FTP_Port;
            vSplit[7] = "infoftp";                        //FTP_Id;
            vSplit[8] = "Infof12X";                       //FTP_Password;


            vSplit[12] = "전호수"; //LoginDescription;
            vSplit[13] = "전호수(100937)"; //LoginDisplayName;
                        
            vSplit[22] = "570"; //DepartmentID
            vSplit[23] = "정보전략팀"; //DepartmentName
            vSplit[24] = "Flex_ERP\\Kor"; //mBaseWorkingDirectory

            ////SOB - 10
            ////BH.
            //vSplit[0] = "59.16.125.7";                  //Oracle_Host
            //vSplit[1] = "1521";                           //Oracle_Port
            //vSplit[2] = "MESORA";                          //Oracle_ServiceName
            //vSplit[3] = "APPS";                           //Oracle_UserId
            //vSplit[4] = "infoflex";                       //Oracle_Password       

            //vSplit[5] = "211.168.59.132";                 //FTP_Host;
            //vSplit[6] = "21";                             //FTP_Port;
            //vSplit[7] = "flexftp";                        //FTP_Id;
            //vSplit[8] = "infoflex";                       //FTP_Password;

            //vSplit[9] = "10";   //SOB_ID
            //vSplit[10] = "101"; //ORG_ID

            //vSplit[11] = "206"; //LoginId;
            //vSplit[12] = "서인철"; //LoginDescription;
            //vSplit[13] = "서인철(B07022)"; //LoginDisplayName;

            //vSplit[19] = "B07022"; //LoginNo                        
            //vSplit[20] = "269"; //PersonID
            //vSplit[21] = "B07022"; //PersonNumber
            //vSplit[22] = "266"; //DepartmentID
            //vSplit[23] = "재무파트"; //DepartmentName
            //vSplit[24] = "Flex_ERP\\Kor"; //mBaseWorkingDirectory

            vSplit[14] = DateTime.Now.ToShortDateString(); //LoginDate;
            vSplit[15] = DateTime.Now.ToString("HH:mm:ss", null); //LoginTime;
            vSplit[16] = "KOR"; //TerritoryLanguage
            //vSplit[16] = "ENG"; //TerritoryLanguage
            vSplit[17] = "S"; //UserType - 사용자구분(A.기본USER/B.제한된USER/S.시스템USER)
            vSplit[18] = "S"; //UserAuthorityType - 권한구분 (A.별도정의/S.SUPERUSER)

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new APPF0020(vSplit[0], vSplit[1], vSplit[2], vSplit[3], vSplit[4],
                                         vSplit[5], vSplit[6], vSplit[7], vSplit[8],
                                         vSplit[9], vSplit[10],
                                         vSplit[11], vSplit[12], vSplit[13], vSplit[14], vSplit[15],
                                         vSplit[16], vSplit[17], vSplit[18], vSplit[19],
                                         vSplit[20], vSplit[21], vSplit[22], vSplit[23],
                                         vSplit[24]));
        }
    }
}