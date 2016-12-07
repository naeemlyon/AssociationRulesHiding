using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


namespace AR
{
    public partial class Form1 : Form
    {
        public const String DB = "D";
        public String XL_File = Application.StartupPath + "\\AR.xls";
        public const String DB_File = "AR.mdb";
        public const int Item_Offset = 5;
        public const String ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + DB_File;
        public OleDbConnection Con;       
        public String Program_Title = "Novel Architecture";
        public String Chosen_Weight = "SumOf";
        public int Total_Freq_Items;
        public int Single_Freq_Items;

        public Form1()
        {
         InitializeComponent();
         
         if (System.IO.File.Exists(DB_File))
             System.IO.File.Delete(DB_File);
         System.IO.File.Copy("AR_Backup.mdb", DB_File);
         // initiate DB Connection          
         try
         {
             Con = new OleDbConnection(ConnectionString);
             Con.Open();
         }
         catch (Exception ex)
         {
             MessageBox.Show(ex.Message, Program_Title, MessageBoxButtons.OK, MessageBoxIcon.Hand);
             return;
         }
         lbl_Total_DB_Records.Text = Count_Total_DB_Records().ToString();            
         Create_Menus();
         // Populate (Support & Confidence) , (Probability , PS) ListBoxes
         Populate_Thresholds();
         //Weight_Clause = "";
         mnu_Generate_Weights_Click(null, null);
         //Generate_Dataset_Correlation();
        }
                                         
        private void mnu_Apply_Technique_Click(object sender, EventArgs e)
        {            
            OleDbCommand CMD_1 = new OleDbCommand();            
            OleDbDataReader R;
            String[] Char_No = new String[2];
            int SAR_ID;
            String Antec, Conseq ;
            float Measure_Value , Cur_FOC;
            char[] Ar;
            //--------------------------------------------                                                                        

            txt_Display.Text = "Starting Applying Weights (Sum).....";            
            CMD_1.CommandText = "Select * from SAR";
            CMD_1.Connection = Con;
            R = CMD_1.ExecuteReader();
            while (R.Read())
            {
                Antec = R[1].ToString(); // Antecedent
                Conseq = R[2].ToString(); // Consequent
                // Confidence,Leverage,Lift,Conviction,All-Confidence (index [0 - 4]+6)
                Measure_Value = float.Parse(R[lbx_Measure.SelectedIndex + 6].ToString());
                Cur_FOC = float.Parse(R[Item_Offset].ToString());  // current FOC

                SAR_ID = int.Parse(R[0].ToString());
                Ar = Sort_SAR_Items(chk_FOC_SAR_Asc_Desc.Text);            
                // decide which character will be adjusted to lower down the FOC
                // and also whether it would be from antec or from conseq. at which position. ?
                Char_No = Extract_Character_No(Ar, Antec, Conseq);
                //txt_Display.AppendText(Environment.NewLine + SAR_ID.ToString() + " : " + Char_No[0] + " - " + Char_No[1] );            

                //txt_Display.AppendText(SAR_ID + " : " + Antec + "->" + Conseq + " MeasureValue=" + Measure_Value + Environment.NewLine);                 
                Apply_Technique(SAR_ID, Antec, Conseq, int.Parse(Char_No[1]), Measure_Value, Cur_FOC, Char_No[0], lbx_Measure.Text);
                Refresh_SAR(SAR_ID, ""); // Refresh All SAR now..                
            }
            R.Close();
            //txt_Display.Text = "Technique Applied! Successfully..";
        }

        private void Apply_Technique(int SAR_ID, String Antec, String Conseq, int Char_No, float Measure_Value, float Cur_FOC, String A_R, String Measure)
        {
            String SQL_String;
            OleDbCommand CMD_1 = new OleDbCommand();
            OleDbCommand CMD_2 = new OleDbCommand();
            OleDbDataReader R;
            String S , SAR;
            char Ch;
            int i , Total_DB;
            bool Running;
            //--------------------------------------------       
            if (A_R == "Antecedent") SAR = Antec;
            else SAR = Conseq;
            CMD_1.Connection = Con;
            CMD_2.Connection = Con;
            Total_DB = int.Parse(lbl_Total_DB_Records.Text);
            //--------------------------------------------                        
            String[] TranName = new String[Total_DB];
            int[] TID = new int[Total_DB];
            //--------------------------------------------                        

            SQL_String = "select * from " + DB + " order by " + lbx_Weights.Text +   " Desc";
            CMD_1.CommandText = SQL_String;            
            R = CMD_1.ExecuteReader();
            int j = 0;
            while (R.Read())
            {
                // Single Record Complete Transaction in form of ABDEFGK                
                TranName[j] = "";
                for (i = 1; i < (R.FieldCount - Item_Offset); i++)
                {
                    S = R[i].ToString();
                    if (S.Contains("1"))
                    {
                        TranName[j] += R.GetName(i).ToString(); 
                    }
                }
                TID[j++] = int.Parse(R[0].ToString());
            }
            R.Close();
            //--------------------------------------------                               
            i = -1;
            Running = Measure_Threshold(Measure, Measure_Value, Cur_FOC);
            while ((Running) && (++i < Total_DB) ) 
            {
                if (Is_Subset(TranName[i], (Antec + Conseq)))
                {
                    Ch = SAR[Char_No]; // Character of each SAR                    

                    CMD_2.CommandText = "update " + DB + " Set " + Ch.ToString() + "=0 Where ID=" + TID[i].ToString();                    
                    CMD_2.ExecuteNonQuery();
                    // txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine );
                    CMD_2.CommandText = "update " + DB + " Set Modification=0 Where ID=" + TID[i].ToString();
                    CMD_2.ExecuteNonQuery();
            
                    Measure_Value = Refresh_SAR(SAR_ID, lbx_Measure.Text ); // Refresh all SAR                                         
                    Running = Measure_Threshold(Measure, Measure_Value, Cur_FOC);
                    lbl_Modified_DB_Records.Text = Count_Modified_DB_Records().ToString();
                    lbl_Modified_DB_Records.Refresh();
                    //Weight_Clause = " Where ID=" + TID[i].ToString();
                    
                }
            }
        }

        private bool Measure_Threshold(String Measure, float Measure_Value, float Cur_FOC)
        {
            bool Ret=false;
            switch (Measure)
            {
                case "Confidence":
                    {
                        Ret = (Measure_Value >= float.Parse(lst_Confidence.SelectedItem.ToString()));            
                        break;
                    }
                case "All-Confidence":
                    {
                        Ret = (Measure_Value >= float.Parse(lst_All_Confidence.SelectedItem.ToString()));                        
                        break;
                    }                
                case "Leverage":
                    {                     
                        Ret = (Measure_Value >= float.Parse(tbx_Min_Leverage.Text)); 
                        break;
                    }
                case "Lift":
                    {
                        Ret = (Measure_Value >= float.Parse(tbx_Min_Lift.Text));
                        break;
                    }
                case "Conviction":
                    {
                        Ret = (Measure_Value >= float.Parse(tbx_Min_Conviction.Text));
                        break;
                    }
            }
            Ret = Ret & (Cur_FOC >= float.Parse(lst_Support.SelectedItem.ToString())); // current FOC comparison
            return Ret;
        }

        private void mnu_Generate_Result_Click(object sender, System.EventArgs e)
        {
            int i, j;
            lvw_Report.Items.Clear();            
            mnu_FreqItems_AssocRules_Click(null, null); // Generate Actual Freq Items and Assocaition Rules
            lvw_Report.Tag = Total_Freq_Items; // Actual FreqItems without ApplyTechnique()
            lvw_Backup.Tag = lvw_Freq.Items.Count; // Actual Assoc.Rules without ApplyTechnique()
            mnu_Populate_Click(null, null);            
            grp_Progress.Visible = true;
            for (i = 0; i < lbx_Measure.Items.Count; i++)
            {                
                lbx_Measure.SelectedIndex = i;
                for (j = 0; j < lbx_Weights.Items.Count; j++)
                {
                    lbx_Weights.SelectedIndex = j;                     
                    grpb_Measure.Refresh();
                    mnu_Generate_Weights_Click(null, null);
                    mnu_Apply_Technique_Click(null, null);
                    mnu_FreqItems_AssocRules_Click(null, null);
                    ///////////////////////////////////////////////////////////                    
                    //Con.Close();
                    Con.Dispose();                    
                    Fill_Report();
                    Fill_Progress_Bar();                    

                    MessageBox.Show(lbx_Measure.Text + Environment.NewLine + lbx_Weights.Text, Program_Title, MessageBoxButtons.OK,  MessageBoxIcon.Asterisk ,MessageBoxDefaultButton.Button1);
                    System.IO.File.Copy("AR_Backup.mdb", DB_File , true );                    
                    // initiate DB Connection          
                    try
                    {
                        Con = new OleDbConnection(ConnectionString);
                        Con.Open();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, Program_Title, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        return;                        
                    }                    
                }
            }
            grp_Progress.Visible = false;
        }
        
        private void Fill_Report()
        {
            ListViewItem L = new ListViewItem();
            String S =String.Empty;
            int Cnt;
            int FI_Cnt = int.Parse(lvw_Report.Tag.ToString()); // Actual FreqItems without ApplyTechnique()
            int AR_Cnt = int.Parse(lvw_Backup.Tag.ToString()); // Actual Assoc.Rules without ApplyTechnique()
            //--------------------------------------------               
            Cnt = lvw_Report.Items.Count + 1;
            
            L = lvw_Report.Items.Add(Cnt.ToString());
            L.SubItems.Add(lbx_Measure.Text);
            L.SubItems.Add(lbx_Weights.Text);
            switch (lbx_Measure.Text)
            {
                case "Confidence":
                    {
                        S = lst_Confidence.Text;
                        Filter_out_AR(7, lst_Confidence.Text); 
                        break;
                    }
                case "Leverage":
                    {
                        S = tbx_Min_Leverage.Text;
                        Filter_out_AR(8, float.Parse(tbx_Min_Leverage.Text), float.Parse(tbx_Max_Leverage.Text));
                        break;
                    }
                case "Lift":
                    {
                        S = tbx_Min_Lift.Text;
                        Filter_out_AR(9, float.Parse(tbx_Min_Lift.Text), float.Parse(tbx_Max_Lift.Text));
                        break;
                    }
                case "Conviction":
                    {
                        S = tbx_Min_Conviction.Text;
                        Filter_out_AR(10, float.Parse(tbx_Min_Conviction.Text), float.Parse(tbx_Max_Conviction.Text));   
                        break;
                    }
                case "All-Confidence":
                    {
                        S = lst_All_Confidence.Text;
                        Filter_out_AR(11, lst_All_Confidence.Text);                         
                        break;
                    }
            }
            
            L.SubItems.Add(S); // Threshold            
            L.SubItems.Add(FI_Cnt.ToString()); // Total Frequent Items            
            L.SubItems.Add(AR_Cnt.ToString()); // Total Association Rules 

            FI_Cnt -= Total_Freq_Items;  // Lost Frequent Items
            L.SubItems.Add(FI_Cnt.ToString());

            AR_Cnt -= lvw_Freq.Items.Count; // Lost Association Rules 
            L.SubItems.Add(AR_Cnt.ToString());

            L.SubItems.Add(lbl_Modified_DB_Records.Text); // Modified DB Transactions
            
            lvw_Report.EnsureVisible(lvw_Report.Items.Count-1);

        }
                
    }
}