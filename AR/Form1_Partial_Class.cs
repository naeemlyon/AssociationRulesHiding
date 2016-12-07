using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Collections;
using System.IO;


namespace AR
{
    // Implements the manual sorting of items by columns.
    class ListViewItemComparer : IComparer
    {
        private int col;
        public ListViewItemComparer()
        {
            col = 0;
        }
        public ListViewItemComparer(int column)
        {
            col = column;
        }
        public int Compare(object x, object y)
        {
            return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
        }
    }


    public partial class Form1 : Form
    {
        //public String Weight_Clause;
        public void Create_Menus()
        {
            // Create a main menu object.
            MainMenu mainMenu1 = new MainMenu();
            // Create empty menu item objects.
            MenuItem mnu_Initialize = new MenuItem();
            MenuItem mnu_Populate = new MenuItem();
            MenuItem mnu_Sep_1 = new MenuItem();
            MenuItem mnu_Refresh_DB = new MenuItem();
            MenuItem mnu_SAR = new MenuItem();
            MenuItem mnu_Add_SAR = new MenuItem();
            MenuItem mnu_Delete_SAR = new MenuItem();
            MenuItem mnu_Truncate_SAR = new MenuItem();
            MenuItem mnu_Generate = new MenuItem();
            MenuItem mnu_FreqItemset = new MenuItem();
            MenuItem mnu_AssocRules = new MenuItem();
            MenuItem mnu_FreqItems_AssocRules = new MenuItem();
            MenuItem mnu_Sep_2 = new MenuItem();
            MenuItem mnu_Extract_Measure = new MenuItem();
            MenuItem mnu_Generate_Weights = new MenuItem();
            MenuItem mnu_Apply_Technique = new MenuItem();
            MenuItem mnu_Report = new MenuItem();
            MenuItem mnu_Generate_Result = new MenuItem();
            MenuItem mnu_Excel = new MenuItem();
            MenuItem mnu_Sort = new MenuItem();
            MenuItem mnu_Sort_Support = new MenuItem();
            MenuItem mnu_Sort_Confidence = new MenuItem();
            MenuItem mnu_Sort_Leverage = new MenuItem();
            MenuItem mnu_Sort_Lift = new MenuItem();
            MenuItem mnu_Sort_Conviction = new MenuItem();
            MenuItem mnu_Sort_All_Confidence = new MenuItem();
            MenuItem mnu_Sep_3 = new MenuItem();
            MenuItem mnu_Sort_Ascending = new MenuItem();
            MenuItem mnu_Sort_Descending = new MenuItem();

            // Set the caption of the menu items.
            mnu_Initialize.Text = "&Initialize";
            mnu_Populate.Text = "Populate";
            mnu_Sep_1.Text = "-";
            mnu_Refresh_DB.Text = "Refresh DB";
            mnu_SAR.Text = "&SAR";
            mnu_Add_SAR.Text = "Add";
            mnu_Delete_SAR.Text = "Delete";
            mnu_Truncate_SAR.Text = "Truncate";
            mnu_Generate.Text = "Generate";
            mnu_FreqItemset.Text = "Freq Itemset";
            mnu_AssocRules.Text = "Assoc. Rules ";
            mnu_FreqItems_AssocRules.Text = "FreqItems && AssocRules";
            mnu_Sep_2.Text = "-";
            mnu_Extract_Measure.Text = "Extract Measure";            
            mnu_Generate_Weights.Text = "Generate Weights";
            mnu_Apply_Technique.Text = "&Apply Technique";
            mnu_Report.Text = "&Report";
            mnu_Generate_Result.Text = "Generate Result";
            mnu_Excel.Text = "Excel";
            mnu_Sort.Text = "Sort";
            mnu_Sort_Support.Text = "Support";
            mnu_Sort_Confidence.Text = "Confidence";            
            mnu_Sort_Leverage.Text = "Leverage";
            mnu_Sort_Lift.Text  = "Lift";
            mnu_Sort_Conviction.Text  = "Conviction";
            mnu_Sort_All_Confidence.Text  = "All-Confidence";
            mnu_Sep_3.Text = "-";
            mnu_Sort_Ascending.Text = "Ascending";
            mnu_Sort_Ascending.Checked = true;
            mnu_Sort_Descending.Text = "Descending";

            // Add the menu items to the main menu.
            mnu_Initialize.MenuItems.Add(mnu_Populate);
            mnu_Initialize.MenuItems.Add(mnu_Sep_1);            
            mnu_Initialize.MenuItems.Add(mnu_Generate_Weights);
            mnu_SAR.MenuItems.Add(mnu_Add_SAR);
            mnu_SAR.MenuItems.Add(mnu_Delete_SAR);
            mnu_SAR.MenuItems.Add(mnu_Truncate_SAR);
            mnu_Generate.MenuItems.Add(mnu_FreqItemset);
            mnu_Generate.MenuItems.Add(mnu_AssocRules);
            mnu_Generate.MenuItems.Add(mnu_FreqItems_AssocRules);
            mnu_Generate.MenuItems.Add(mnu_Sep_2);
            mnu_Generate.MenuItems.Add(mnu_Extract_Measure);
            mainMenu1.MenuItems.Add(mnu_Refresh_DB);
            mainMenu1.MenuItems.Add(mnu_Initialize);
            mainMenu1.MenuItems.Add(mnu_SAR);
            mainMenu1.MenuItems.Add(mnu_Generate);
            mainMenu1.MenuItems.Add(mnu_Apply_Technique);
            mainMenu1.MenuItems.Add(mnu_Report);
            mnu_Report.MenuItems.Add(mnu_Generate_Result);
            mnu_Report.MenuItems.Add(mnu_Excel);
            mainMenu1.MenuItems.Add(mnu_Sort);
            mnu_Sort.MenuItems.Add(mnu_Sort_Support);
            mnu_Sort.MenuItems.Add(mnu_Sort_Confidence);
            mnu_Sort.MenuItems.Add(mnu_Sort_Leverage);
            mnu_Sort.MenuItems.Add(mnu_Sort_Lift);
            mnu_Sort.MenuItems.Add(mnu_Sort_Conviction);
            mnu_Sort.MenuItems.Add(mnu_Sort_All_Confidence);
            mnu_Sort.MenuItems.Add(mnu_Sep_3);
            mnu_Sort.MenuItems.Add(mnu_Sort_Ascending);
            mnu_Sort.MenuItems.Add(mnu_Sort_Descending); 
            
            // Add functionality to the menu items using the Click event.                         
            mnu_Populate.Click += new System.EventHandler(this.mnu_Populate_Click);
            mnu_Refresh_DB.Click += new System.EventHandler(this.mnu_Refresh_DB_Click);
            mnu_Add_SAR.Click += new System.EventHandler(this.mnu_Add_SAR_Click);
            mnu_Delete_SAR.Click += new System.EventHandler(this.mnu_Delete_SAR_Click);
            mnu_Truncate_SAR.Click += new System.EventHandler(this.mnu_Truncate_SAR_Click);
            mnu_FreqItemset.Click += new System.EventHandler(this.mnu_FreqItemset_Click);
            mnu_AssocRules.Click += new System.EventHandler(this.mnu_AssocRules_Click);
            mnu_FreqItems_AssocRules.Click += new System.EventHandler(this.mnu_FreqItems_AssocRules_Click);
            mnu_Extract_Measure.Click += new System.EventHandler(this.mnu_Extract_Measure_Click);            
            mnu_Generate_Weights.Click += new System.EventHandler(this.mnu_Generate_Weights_Click);
            mnu_Apply_Technique.Click += new System.EventHandler(this.mnu_Apply_Technique_Click);
            mnu_Generate_Result.Click += new System.EventHandler(this.mnu_Generate_Result_Click);
            mnu_Excel.Click += new System.EventHandler(this.mnu_Excel_Click);
            mnu_Sort_Support.Click += new System.EventHandler(this.mnu_Sort_Click);
            mnu_Sort_Confidence.Click += new System.EventHandler(this.mnu_Sort_Click);
            mnu_Sort_Leverage.Click += new System.EventHandler(this.mnu_Sort_Click);
            mnu_Sort_Lift.Click += new System.EventHandler(this.mnu_Sort_Click);
            mnu_Sort_Conviction.Click += new System.EventHandler(this.mnu_Sort_Click);
            mnu_Sort_All_Confidence.Click += new System.EventHandler(this.mnu_Sort_Click);
            mnu_Sort_Ascending.Click += new System.EventHandler(this.mnu_Sort_Click);
            mnu_Sort_Descending.Click += new System.EventHandler(this.mnu_Sort_Click);

            // Assign mainMenu1 to the form.
            this.Menu = mainMenu1;
        }

        private void Populate_Thresholds()
        {
            for (int i = 1; i < 101; i++)
            {
                lst_Support.Items.Add(i);                
                lst_Confidence.Items.Add(i);
                lst_All_Confidence.Items.Add(i);
            }
            lst_Support.SelectedIndex = 49;            
            lst_Confidence.SelectedIndex = 64;
            lst_All_Confidence.SelectedIndex = 49;
            lbx_Measure.SelectedIndex = 0;
            lbx_Weights.SelectedIndex = 0;
            lbl_Progress.BackColor = pbr_Busy.BackColor;
            grp_Progress.Visible = false;
            lbl_Total_Items.Text = Count_Total_DB_Items().ToString();
        }
                
        // this function no more used in this program but valuable..
        private String Merge_Only_Unique_Characters(String S1, String S2)
        {
            //example: ACGF + ACD = ACDF (A&C only one time..)
            int i = 0, j = 0, k = 0, l = 0; String Ret = "";
            bool Unique = false;
            for (i = 0; i < S1.Length; i++) // S1 will be fixed
            {
                for (j = 0; j < S2.Length; j++) // S2 will be extracted with only unique characters
                {
                    if (S1[i] == S2[j]) // if any Same Character
                    {
                        for (k = 0; k < S2.Length; k++) // Add up all characters of S2 to S1 but only unique characters...
                        {
                            Unique = true;
                            for (l = 0; l < S1.Length; l++)
                            {
                                if (S1[l] == S2[k])
                                {
                                    Unique = false;
                                }
                            }
                            if (Unique == true) Ret += S2[k];
                        }
                        Ret = S1 + Ret;
                        i = S1.Length + 1;
                        j = S2.Length + 1;
                    }
                }
            }
            return (Ret);
        }

        private void Generate_Dataset_Correlation()
        {
            String SQL_String;
            OleDbCommand CMD = new OleDbCommand();
            OleDbCommand C1 = new OleDbCommand();
            OleDbDataReader R;
            ListViewItem L = new ListViewItem();
            String S;            
            int i, Sup, Total = 0, Cntr = 0; //,j;
            //--------------------------------------------             
            SQL_String = "select * from D";
            CMD.CommandText = SQL_String;
            CMD.Connection = Con;
            R = CMD.ExecuteReader();
            lvw_Freq.Items.Clear();
            for (i = 1; i < (R.FieldCount - Item_Offset); i++)
            {
                S = R.GetName(i).ToString();
                Sup = Calculate_Frequency(S);
                /* // only display no. of items in listview..
                L = lvw_Freq.Items.Add(S);
                j = 6;
                while (--j > 0) L.SubItems.Add("");
                L.SubItems.Add(Sup.ToString());
                */
                Total += Sup;
                Cntr++;
            }
            R.Close();
            float Correlation_Factor = (float.Parse(Total.ToString()) / Cntr);
            Correlation_Factor = Correlation_Factor / int.Parse(lbl_Total_DB_Records.Text);
            Correlation_Factor *= 100;
            lbl_Total_DB_Records.Tag = Correlation_Factor;
            this.Text = Correlation_Factor.ToString();
        }

        private int Count_Total_DB_Records()
        {
            OleDbCommand Cmd = new OleDbCommand();
            int Ret;
            Cmd.CommandText = "SELECT count(*) from D ";
            Cmd.Connection = Con;
            Ret = (int)Cmd.ExecuteScalar();
            return Ret;
        }

        private int Count_Modified_DB_Records()
        {
            OleDbCommand Cmd = new OleDbCommand();
            int Ret;
            Cmd.CommandText = "SELECT count(*) from D Where Modification=0";
            Cmd.Connection = Con;
            Ret = (int)Cmd.ExecuteScalar();
            return Ret;
        }

        public void Load_Data_Grid(string sqlQueryString, DataGridView DGV)
        {
            OleDbCommand SQLQuery = new OleDbCommand();
            DataTable data = null;
            DGV.DataSource = null;
            SQLQuery.Connection = null;
            OleDbDataAdapter dataAdapter = null;
            DGV.Columns.Clear();
            //---------------------------------
            SQLQuery.CommandText = sqlQueryString;
            SQLQuery.Connection = Con;
            data = new DataTable();
            dataAdapter = new OleDbDataAdapter(SQLQuery);
            dataAdapter.Fill(data);
            DGV.DataSource = data;
            DGV.AllowUserToAddRows = false; // remove the null line
            DGV.ReadOnly = true;
            DGV.Columns[0].Visible = true;
            /*DGV.Columns[0].Width = 20;
            
            foreach (ColumnHeader ch in this.lvw_Freq.Columns)
            {
                ch.Width = -2;
            }
            */
            DGV.Refresh();
        }

        private int Count_Total_DB_Items()
        {
            OleDbCommand CMD_1 = new OleDbCommand();            
            OleDbDataReader R;            
            CMD_1.CommandText = "select * from " + DB ;
            CMD_1.Connection = Con;
            R = CMD_1.ExecuteReader();            
            int Ret = (R.FieldCount - Item_Offset);
            R.Dispose(); 
            CMD_1.Dispose();
            return Ret;
        }

        private void mnu_Sort_Click(object sender, EventArgs e)
        {
            int Sort_Index = 0;
            // move below at class level. 
            // public SortOrder Sort_Order_Measure;
            //((MenuItem)sender).Checked = true; // unckek other too

            if (((MenuItem)sender).Text == "Ascending")
            {
                //Sort_Order_Measure = SortOrder.Ascending;
             
                MessageBox.Show("Yet to implement");
                return;
            }
            else if (((MenuItem)sender).Text == "Descending")
            {
                //Sort_Order_Measure = SortOrder.Descending;                
                MessageBox.Show("Yet to implement");
                return;
            }

            if (((MenuItem)sender).Text == "Support")
                Sort_Index = 6;
            else if (((MenuItem)sender).Text == "Confidence")
                Sort_Index = 7;
            else if (((MenuItem)sender).Text == "Leverage")
                Sort_Index = 8;
            else if (((MenuItem)sender).Text == "Lift")
                Sort_Index = 9;
            else if (((MenuItem)sender).Text == "Conviction")
                Sort_Index = 10;
            else if (((MenuItem)sender).Text == "All-Confidence")
                Sort_Index = 11;
            else
            {
                lvw_Freq.Sorting = SortOrder.None;
                lvw_Freq.ListViewItemSorter = new ListViewItemComparer(0);
                return;
            }
            //lvw_Freq.Sorting = Sort_Order_Measure;
            lvw_Freq.Sorting = SortOrder.Ascending;
            lvw_Freq.ListViewItemSorter = new ListViewItemComparer(Sort_Index);            
            lvw_Freq.Sort();            
        }

        private void Colorize_List_View()
        {
            int i;
            float Min_PS = 100; float Max_PS = -100;
            for (i = 0; i < lvw_Freq.Items.Count; i++)
            {
                lvw_Freq.Items[i].UseItemStyleForSubItems = false;
                lvw_Freq.Items[i].SubItems[1].BackColor = Color.Azure;  // Anctecdent (Tail) 
                lvw_Freq.Items[i].SubItems[1].ForeColor = Color.Chocolate; //Antecedent (Tail) 
                lvw_Freq.Items[i].SubItems[3].BackColor = Color.Azure; // Consequent (Tail) 
                lvw_Freq.Items[i].SubItems[3].ForeColor = Color.Chocolate; //Consequent (Tail) 
                lvw_Freq.Items[i].SubItems[7].BackColor = Color.LightGoldenrodYellow; // Confidence
                lvw_Freq.Items[i].SubItems[7].ForeColor = Color.BlueViolet;   // Confidence 
                lvw_Freq.Items[i].SubItems[8].BackColor = Color.LightGray; // Leverage(PS)
                lvw_Freq.Items[i].SubItems[8].ForeColor = Color.Red;   // Leverage (PS)
                lvw_Freq.Items[i].SubItems[9].BackColor = Color.Wheat; //Lift
                lvw_Freq.Items[i].SubItems[9].ForeColor = Color.Blue;   // Lift
                lvw_Freq.Items[i].SubItems[10].BackColor = Color.Gray;  //Conviction 
                lvw_Freq.Items[i].SubItems[10].ForeColor = Color.Gold;       // Conviction 
                lvw_Freq.Items[i].SubItems[11].BackColor = Color.MistyRose;  //All-Confidenc 
                lvw_Freq.Items[i].SubItems[11].ForeColor = Color.Green;       // All-Confidence 

                if (System.Convert.ToDouble(lvw_Freq.Items[i].SubItems[8].Text) > Max_PS)
                {
                    Max_PS = float.Parse(lvw_Freq.Items[i].SubItems[8].Text);
                }
                if (Min_PS > System.Convert.ToDouble(lvw_Freq.Items[i].SubItems[8].Text))
                {
                    Min_PS = float.Parse(lvw_Freq.Items[i].SubItems[8].Text);
                }
            }
            lvw_Freq.Refresh();
        }

        private void Copy_ListView_Data(ListView Src, ListView Dest)
        {
            int i, j;
            ListViewItem L = new ListViewItem();
            Dest.Items.Clear();
            for (i = 0; i < Src.Items.Count; i++)
            {
                L = Dest.Items.Add(Src.Items[i].Text);
                for (j = 1; j < Src.Items[i].SubItems.Count; j++)
                {
                    L.SubItems.Add(Src.Items[i].SubItems[j].Text);
                }
            }
        }

        private void mnu_Refresh_DB_Click(object sender, EventArgs e)
        {
            Con.Dispose();            
            MessageBox.Show("Origional Data Imported", Program_Title, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            System.IO.File.Copy("AR_Backup.mdb", DB_File, true);
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

        private void mnu_Truncate_SAR_Click(object sender, EventArgs e)
        {
            OleDbCommand CMD = new OleDbCommand();
            DialogResult Ans = MessageBox.Show("All SAR will be erased from database!" + Environment.NewLine + "Proceed ?", Program_Title, MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
            if (Ans != DialogResult.Yes) return;
            CMD.CommandText = "Delete from SAR";
            CMD.Connection = Con;
            CMD.ExecuteNonQuery();
        }

        private void mnu_Add_SAR_Click(object sender, EventArgs e)
        {
            int Indx;
            String SQL;
            OleDbCommand CMD = new OleDbCommand();
            //////////////////////////////////////////////
            if (lvw_Freq.SelectedItems.Count == 0) return;
            Indx = lvw_Freq.SelectedItems[0].Index;

            SQL = "INSERT into SAR (Antecedent, Consequent, LSup, RSup, Support, Confidence, Leverage, Lift, Conviction, All_Confidence) VALUES(";
            SQL += "'" + lvw_Freq.Items[Indx].SubItems[1].Text + "', ";  // Antecedent
            SQL += "'" + lvw_Freq.Items[Indx].SubItems[3].Text + "', ";  // Consequent
            SQL += lvw_Freq.Items[Indx].SubItems[4].Text + " , ";  // LSup
            SQL += lvw_Freq.Items[Indx].SubItems[5].Text + " , ";  // RSup
            SQL += lvw_Freq.Items[Indx].SubItems[6].Text + " , ";  // Support
            SQL += lvw_Freq.Items[Indx].SubItems[7].Text + " , ";  // Confidence
            SQL += lvw_Freq.Items[Indx].SubItems[8].Text + " , ";  // Leverage
            SQL += lvw_Freq.Items[Indx].SubItems[9].Text + " , ";  // Lift
            SQL += lvw_Freq.Items[Indx].SubItems[10].Text + " , "; // Conviction
            SQL += lvw_Freq.Items[Indx].SubItems[11].Text ;  // All_Confidence
            SQL += ")";

            SQL = SQL.Replace("Infinity", "1000");
            CMD.CommandText = SQL;            
            CMD.Connection = Con;
            CMD.ExecuteNonQuery();
            txt_Display.Text = "Selected Assoc. Rule inserted into SAR set successfully...";
        }

        private void mnu_Delete_SAR_Click(object sender, EventArgs e)
        {
            String SQL;
            OleDbCommand CMD = new OleDbCommand();
            //////////////////////////////////////////////
            if (dgv_SAR.SelectedRows.Count == 0) return;
            SQL = "Delete from SAR WHERE ID=" + dgv_SAR.SelectedRows[0].Cells[0].Value;
            DialogResult Ans = MessageBox.Show("Do U really want to delete Selected SAR", Program_Title, MessageBoxButtons.YesNo, MessageBoxIcon.Hand);
            if (Ans != DialogResult.Yes) return;
            CMD.CommandText = SQL;
            CMD.Connection = Con;
            CMD.ExecuteNonQuery();
            txt_Display.Text = "Selected SAR Deleted Successfully...Refresh Grid View";
        }

        private void mnu_Populate_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (chk_dgv_Data.Checked == true)
             Load_Data_Grid("SELECT * FROM D", dgv_Data);
            if (chk_dgv_SAR.Checked == true)
             Load_Data_Grid("SELECT * FROM SAR", dgv_SAR);
            Cursor.Current = Cursors.Default;
        }

        private void mnu_FreqItemset_Click(object sender, EventArgs e)
        {            
            Generate_Single_Item_Freq_Table();
            //MessageBox.Show("");
            Single_Freq_Items = lvw_Freq.Items.Count;             
            Generate_Frequent_Items();
            Remove_Duplicate_Freq_Items(lvw_Freq);
            Total_Freq_Items = Single_Freq_Items + lvw_Freq.Items.Count;
            txt_Display.Text = Single_Freq_Items + " Single Freq Items and " + (lvw_Freq.Items.Count).ToString() + " Multiple value Frequent Items Generated:  Total = " + Total_Freq_Items;
        }

        private void mnu_AssocRules_Click(object sender, EventArgs e)
        {
            Generate_Association_Rules();
            Colorize_List_View();
            Copy_ListView_Data(lvw_Freq, lvw_Backup);                        
        }

        private void mnu_FreqItems_AssocRules_Click(object sender, EventArgs e)
        {            
            mnu_FreqItemset_Click(null, null);
            mnu_AssocRules_Click(null, null);
            Generate_Tab_Text(lvw_Freq);
            //mnu_Extract_Measure_Click(null, null);
        }

        private void Generate_Single_Item_Freq_Table()
        {
            String SQL_String;
            OleDbCommand CMD = new OleDbCommand();
            OleDbCommand C1 = new OleDbCommand();
            OleDbDataReader R;
            ListViewItem L = new ListViewItem();
            String S;
            float Given_Sup;
            int i, j, Sup;
            //--------------------------------------------             
            SQL_String = "select * from D";
            CMD.CommandText = SQL_String;
            CMD.Connection = Con;
            R = CMD.ExecuteReader();
            lvw_Freq.Items.Clear();
            for (i = 1; i < (R.FieldCount - Item_Offset); i++)
            {
                S = R.GetName(i).ToString();
                Sup = Calculate_Frequency(S);
                Given_Sup = float.Parse(lst_Support.Text);
                Given_Sup = (Given_Sup / 100) * int.Parse(lbl_Total_DB_Records.Text);
                if (Sup >= Given_Sup)
                {
                    L = lvw_Freq.Items.Add(S);
                    j = 6;
                    while (--j > 0) L.SubItems.Add("");
                    L.SubItems.Add(Sup.ToString());
                }
            }
            R.Close();
        }

        private void Write_Frequent_Items(int[] mask, int n, String Str)
        {
            int i;
            String S = "";
            ListViewItem L = new ListViewItem();
            int Sup;
            float Given_Sup;

            for (i = 0; i < n; ++i)
            {
                if (mask[i] > 0)
                {
                    S += Str[i].ToString();
                }
            }
            if (S.Length > 1)
            {
                Sup = Calculate_Frequency(S);

                Given_Sup = float.Parse(lst_Support.Text);
                Given_Sup = (Given_Sup / 100) * int.Parse(lbl_Total_DB_Records.Text);
                if (Sup >= Given_Sup)
                {
                    L = lvw_Freq.Items.Add(S);
                    for (i = 1; i < 6; i++) L.SubItems.Add("");
                    L.SubItems.Add(Sup.ToString());
                }
            }
        }

        private int Get_Next_Frequent_Items(int[] mask, int n)
        {
            int i;
            for (i = 0; ((i < n) && (mask[i] > 0)); ++i)
            {
                mask[i] = 0;
            }

            if (i < n)
            {
                mask[i] = 1;
                return 1;
            }
            return 0;
        }

        private void Generate_Frequent_Items()
        {
            int i, n, UL;
            String S = "";
            int[] mask = new int[20]; /* Guess what this is */

            UL = lvw_Freq.Items.Count;
            for (i = 0; i < UL; i++) // traverse each and every frequent-item
            {
                S += lvw_Freq.Items[i].Text;
            }
            n = S.Length;
            for (i = 0; i < n; ++i)
            {
                mask[i] = 0;
            }
            // Get all the other sub sets excluding empty set..
            txt_Display.Text = "Generating Frequent Item...";
            txt_Display.Refresh();
            while (Get_Next_Frequent_Items(mask, n) > 0)
            {
                Write_Frequent_Items(mask, n, S);
                lvw_Freq.Refresh();
            }
            // clean up all the previous itemset now..
            for (i = UL - 1; i >= 0; i--)
            {
                lvw_Freq.Items[i].Remove();
            }                        
        }

        private int Calculate_Frequency(String Str)
        {
            int i;
            String S = "";            
            OleDbCommand CMD = new OleDbCommand();
            OleDbCommand C1 = new OleDbCommand();
            int Freq = 0;
            //----------------------------------------------
            Str = Str.Trim();
            for (i = 0; i < Str.Length; i++)
            {
                S += Str[i] + "=1 AND ";
            }
            if (S.Length > 3)
            {
                S = S.Substring(0, S.Length - 4);                
                C1.CommandText = "SELECT count(*) from " + DB + " where " + S;
                C1.Connection = Con;
                Freq = (int)C1.ExecuteScalar();
            }
            return Freq;
        }
        
        private void Remove_Duplicate_Freq_Items(ListView LVW)
        {
            //-- ACG , AGC , CAG are same so only first ACG should exit. others be removed.
            int i, j, m;
            String S1, S2;
            bool Found = false;

            for (i = 0; i < LVW.Items.Count; i++)
            {
                for (j = LVW.Items.Count - 1; j > i; j--)
                {
                    if (LVW.Items[i].Text.Length == LVW.Items[j].Text.Length)
                    {
                        if (LVW.Items[i].Text == LVW.Items[j].Text)
                        {
                            LVW.Items[j].Remove();
                        }
                        //-- check ACG and AGC are same but above is omitted 
                        //-- so need to be implemented
                        else
                        {
                            S1 = LVW.Items[i].Text;
                            S2 = LVW.Items[j].Text;
                            Found = true; // assume that both are same...
                            for (m = 0; m < S2.Length; m++)
                            {
                                if (S1.Contains(S2[m].ToString()) == false)
                                {
                                    Found = false;
                                }
                            }
                            if (Found == true)
                            {
                                LVW.Items[j].Remove();
                            }
                        }
                    }
                }
            }
        }

        private void Write_Asscociation_Rules(int[] mask, int n, String Str, int j)
        {
            int i;
            ListViewItem L = new ListViewItem();
            float Confidence, PS, Lift, a, c; ;
            int Ant_Sup, Con_Sup, Total_DB;
            String Ant = "", Con = "", S;
            S = Str;
            for (i = 0; i < n; ++i)
            {
                if (mask[i] > 0)
                {
                    Ant += S[i].ToString();
                }
            }
            for (i = 0; i < Ant.Length; i++)
            {
                S = S.Replace(Ant[i].ToString(), ""); // remove each character of Ant from S, 
            }
            Con = S; // now the trimmed S becomes Conequent.                        

            if ((Ant.Length < Str.Length) && (Con.Length < Str.Length))
            {
                Total_DB = int.Parse(lbl_Total_DB_Records.Text);
                Ant_Sup = Calculate_Frequency(Ant);
                Con_Sup = Calculate_Frequency(Con);
                Confidence = (float.Parse(lvw_Freq.Items[j].SubItems[6].Text) / Ant_Sup) * 100;
                L = lvw_Freq.Items.Add(Str);
                L.SubItems.Add(Ant);
                L.SubItems.Add("->");
                L.SubItems.Add(Con);
                a = float.Parse(Ant_Sup.ToString());
                a = (a / Total_DB);
                c = float.Parse(Con_Sup.ToString());
                c = (c / Total_DB);
                float ac = float.Parse(lvw_Freq.Items[j].SubItems[6].Text) / Total_DB;
                PS = ac - (a * c); // Leverage
                a = a * 100;
                L.SubItems.Add(a.ToString());  //  L.SubItems.Add(Ant_Sup.ToString());                
                c = c * 100;
                L.SubItems.Add(c.ToString()); //  L.SubItems.Add(Con_Sup.ToString());                
                ac = ac * 100;
                L.SubItems.Add(ac.ToString());
                L.SubItems.Add(Confidence.ToString());

                // Leverage (PS).............                
                //PS = PS * 100;
                L.SubItems.Add(PS.ToString());

                // Lift
                Lift = Confidence / float.Parse(Con_Sup.ToString());
                Lift = (Total_DB * Lift) / 100;
                L.SubItems.Add(Lift.ToString());

                // Conviction
                float Convic;
                a = (1 - float.Parse(Con_Sup.ToString()) / float.Parse(Total_DB.ToString()));
                c = (1 - (Confidence / 100));
                Convic = a / c;
                L.SubItems.Add(Convic.ToString());

                //All-Confidenc
                float All_Confidence = Measure_All_Confidence(Str);
                L.SubItems.Add(All_Confidence.ToString());

                // interactive....
                L.EnsureVisible();
            }
        }

        private int Get_Next_Association_Rule(int[] mask, int n)
        {
            int i;
            for (i = 0; ((i < n) && (mask[i] > 0)); ++i)
            {
                mask[i] = 0;
            }

            if (i < n)
            {
                mask[i] = 1;
                return 1;
            }
            return 0;
        }

        private void Generate_Association_Rules()
        {
            int i, j, n, UL;
            String S;
            int[] mask = new int[20]; /* Guess what this is */

            UL = lvw_Freq.Items.Count;
            for (j = 0; j < UL; j++) // traverse each and every frequent-item
            {
                S = lvw_Freq.Items[j].Text;
                n = S.Length;
                for (i = 0; i < n; ++i)
                {
                    mask[i] = 0;
                }
                // all the other sub sets 
                while (Get_Next_Association_Rule(mask, n) > 0)
                {
                    Write_Asscociation_Rules(mask, n, S, j);
                    txt_Display.Text = "Generating Assoc.Rules [Processing Multi-Val Frequent Item: " + j.ToString() + " / " + UL.ToString() + "]";
                    lvw_Freq.Refresh();
                    txt_Display.Refresh();
                }
            }
            // clean up all the Frequent Itemset now..
            for (i = UL - 1; i >= 0; i--)
            {
                lvw_Freq.Items[i].Remove();
            }
            //--------------------                       
            txt_Display.Text = lvw_Freq.Items.Count.ToString() + " Association Rules Generated from " + Total_Freq_Items.ToString() + " Freq.Itemsets";
        }
              
        private void mnu_Generate_Weights_Click(object sender, EventArgs e)
        {
            String SQL_String;
            OleDbCommand CMD = new OleDbCommand();
            OleDbDataReader R;
            String S, TranName;
            int i, Total_SAR;
            //--------------------------------------------             

            CMD.CommandText = "Select count(*) from SAR";
            CMD.Connection = Con;
            Total_SAR = (int)CMD.ExecuteScalar();
            String[] SAR = new String[Total_SAR];
            //--------------------------------------------
            SQL_String = "select * from SAR";
            CMD.CommandText = SQL_String;
            CMD.Connection = Con;
            R = CMD.ExecuteReader();
            i = 0;
            while (R.Read())
            {
                SAR[i++] = (R[1].ToString() + R[2].ToString());
            }
            R.Close();
            ////////////////////////////////////////////

            SQL_String = "select * from D";  //+ Weight_Clause ;
            CMD.CommandText = SQL_String;
            CMD.Connection = Con;
            R = CMD.ExecuteReader();
            txt_Display.Text = "Generating Weights.....";
            while (R.Read())
            {
                TranName = "";
                for (i = 1; i < (R.FieldCount - Item_Offset); i++)
                {
                    S = R[i].ToString();
                    if (S.Contains("1"))
                    {
                        TranName += R.GetName(i).ToString();
                    }
                }
                Apply_Weights_MMMS(long.Parse(R[0].ToString()), TranName, SAR);
            }
            R.Close();
            //Weight_Clause = "";
            txt_Display.Text = "Weights Generated Successfully..";
        }

        private void Apply_Weights_MMMS(long TId, String Tran, String[] SAR)
        {
            int i, j, tmp, C = 0;
            int[] Weights = new int[20];
            int Sum = 0;
            float Mean = 0, Median = 0;
            String S, Found_SAR = "";
            OleDbCommand CMD = new OleDbCommand();
            int SZ_SAR = SAR.GetUpperBound(0) * SAR.GetUpperBound(0);
            int[] MD = new int[SZ_SAR];
            int Max_Count = 0;
            String Mode = "";
            //////////////////////////////////////  
            for (i = 0; i <= SAR.GetUpperBound(0); i++)
            {
                S = "";
                for (j = 0; j < SAR[i].Length; j++)
                {
                    if (Tran.Contains(SAR[i][j].ToString()))
                    {
                        S += "1";
                    }
                }
                if (S.Length == SAR[i].Length)
                {
                    Found_SAR += SAR[i];
                }
            }
            //  txt_Display.AppendText(" => " + Found_SAR + " (" );
            ////////////////////////////////////////////

            for (i = 0; i < Weights.Length; i++)
            {
                Weights[i] = 1;
            }
            // count how many indiv item found (A=2 , B=3 , C=1 etc)
            for (i = 0; i < Found_SAR.Length; i++)
            {
                S = Found_SAR[i].ToString();
                for (j = i + 1; j < Found_SAR.Length; j++)
                {
                    if (S == Found_SAR[j].ToString())
                    {
                        Weights[C]++;
                    }
                }
                Found_SAR = Found_SAR.Replace(S, "");
                i--;
                //    txt_Display.AppendText(S + "=" + Weights[C].ToString() + " , ");
                C++;
            }
            // Sum of the weights (Sum =6 [A=2,B=3,C=1])
            for (i = 0; i < C; i++)
            {
                Sum += Weights[i];
            }


            if (Sum == 0) return; // at least one (shortest) SAR(A->F) is associated with..

            // Mean of the weights  
            Mean = float.Parse(Sum.ToString()) / float.Parse((C).ToString());

            /// Bubble Sort all weights...(Asc Order)         

            for (i = 0; i < C - 1; i++)
            {
                for (j = i + 1; j < C; j++)
                {
                    if (Weights[i] > Weights[j])
                    {
                        tmp = Weights[i];
                        Weights[i] = Weights[j];
                        Weights[j] = tmp;
                    }
                }
            }

            // Median of Weights                        
            if (i % 2 == 0) // index from 0 so odd no.
            {
                Median = Weights[(i / 2)]; // middle index out of odd no.
            }
            else // even no. then avg (middle two values.) 
            {
                Median = Weights[(i / 2)] + Weights[(i / 2) + 1];
                Median = Median / 2;
            }

            // Mode of Weights
            for (i = 0; i < MD.GetUpperBound(0); i++)
            {
                MD[i] = 0;
            }
            for (i = 0; i < C; i++)
            {
                MD[Weights[i]]++;
                if (MD[Weights[i]] > Max_Count)
                    Max_Count = MD[Weights[i]];
            }
            for (i = 0; i < SZ_SAR; i++)
            {
                if (MD[i] >= Max_Count)
                    Mode += i + ",";
            }
            if (Mode.Length > 1) // remove last extra , if any
                Mode = Mode.Substring(0, (Mode.Length - 1));

            // sent them all into databse
            CMD.CommandText = "UPDATE " + DB + " SET Mean=" + Mean + ", Mode='" + Mode + "', Median=" + Median + ", SumOf=" + Sum + " WHERE ID=" + TId + "";
            CMD.Connection = Con;
            CMD.ExecuteNonQuery();
            /* 
             txt_Display.AppendText(") Sum=" + Sum.ToString());
             txt_Display.AppendText(" Mean=" + Mean.ToString());
             txt_Display.AppendText(" Median=" + Median.ToString());
             txt_Display.AppendText(" Mode=" + Mode);
             txt_Display.AppendText(Environment.NewLine);                        
             */
        }

        private float Measure_All_Confidence(String S)
        {
            //all-confidence(Z) = supp(Z) / max(support(z element of Z)) = P(Z) / max(P(z element of Z)) 
            float Ret, Z;
            int i, Max_z;
            Ret = 0;
            Z = float.Parse(Calculate_Frequency(S).ToString());
            for (i = 0; i < S.Length; i++)
            {
                Max_z = Calculate_Frequency(S[i].ToString());
                if (float.Parse(Max_z.ToString()) > Ret)
                {
                    Ret = Max_z;
                }
            }
            return (100 * (Z / Ret));
        }

        private void Extract_Measure_Values(int Cnt, TextBox Tbx_Min, TextBox Tbx_Max)
        {
            int i;
            float Min_Val = 100; float Max_Val = -100;
            for (i = 0; i < lvw_Freq.Items.Count; i++)
            {
                if (System.Convert.ToDouble(lvw_Freq.Items[i].SubItems[Cnt].Text) > Max_Val)
                {
                    Max_Val = float.Parse(lvw_Freq.Items[i].SubItems[Cnt].Text);
                }
                if (Min_Val > System.Convert.ToDouble(lvw_Freq.Items[i].SubItems[Cnt].Text))
                {
                    Min_Val = float.Parse(lvw_Freq.Items[i].SubItems[Cnt].Text);
                }
            }
            Tbx_Min.Text = Min_Val.ToString();
            Tbx_Max.Text = Max_Val.ToString();
        }

        private void Extract_Measure_Values(int Cnt, ListBox LBx)
        {
            int i;
            float Min_Val = 100;
            for (i = 0; i < lvw_Freq.Items.Count; i++)
            {
                if (System.Convert.ToDouble(lvw_Freq.Items[i].SubItems[Cnt].Text) < Min_Val)
                {
                    Min_Val = float.Parse(lvw_Freq.Items[i].SubItems[Cnt].Text);
                }              
            }
            i = (int)Min_Val;
            LBx.Text = (i.ToString()); 
        }

        private void mnu_Extract_Measure_Click(object sender, EventArgs e)
        {
            Extract_Measure_Values(7, lst_Confidence);
            Extract_Measure_Values(8, tbx_Min_Leverage, tbx_Max_Leverage);
            Extract_Measure_Values(9, tbx_Min_Lift, tbx_Max_Lift);
            Extract_Measure_Values(10, tbx_Min_Conviction, tbx_Max_Conviction);
            Extract_Measure_Values(11, lst_All_Confidence);
        }
        
        private void rad_Confidence_Click(object sender, EventArgs e)
        {
            // clean up assocaition rules with below confidence             
            if (((RadioButton)sender).Text == "Confidence %")
            {                
                Filter_out_AR(7, lst_Confidence.Text); 
            }
            else if (((RadioButton)sender).Text == "All-Confidence %")
            {
               Filter_out_AR(11,lst_All_Confidence.Text); 
            }
        }
   
        private void Filter_out_AR(int j, String Confd)
        {
            int i;
            Copy_ListView_Data(lvw_Backup, lvw_Freq);
            for (i = (lvw_Freq.Items.Count - 1); i > 0; i--)
            {
                if (float.Parse(lvw_Freq.Items[i].SubItems[j].Text) < float.Parse(Confd))
                {
                    lvw_Freq.Items[i].EnsureVisible();
                    lvw_Freq.Items[i].Remove();
                    lvw_Freq.Refresh();
                    txt_Display.Text = "Processing Assoc.Rule # " + i.ToString();
                    txt_Display.Refresh();
                }
            }
            Colorize_List_View();
            txt_Display.Text = (lvw_Freq.Items.Count).ToString() + " Association Rules Generated " ; //:  Measure:" + ((RadioButton)sender).Text;
        }

        private void rad_Measure_Click(object sender, EventArgs e)
        {
            // clean up assocaition rules with below confidence             
            if (lvw_Freq.Items.Count == 0) return;
            
            if (((RadioButton)sender).Text == "Leverage (PS)")
            {
              Filter_out_AR(8, float.Parse(tbx_Min_Leverage.Text), float.Parse(tbx_Max_Leverage.Text));
            }
            else if (((RadioButton)sender).Text == "Lift")
            {
              Filter_out_AR(9, float.Parse(tbx_Min_Lift.Text), float.Parse(tbx_Max_Lift.Text));
            }
            else if (((RadioButton)sender).Text == "Conviction")
            {
              Filter_out_AR(10, float.Parse(tbx_Min_Conviction.Text) , float.Parse(tbx_Max_Conviction.Text));
            }
        }

        private void Filter_out_AR(int j, float Min_Val, float Max_Val)
        {
            int i;
            String S;

            Copy_ListView_Data(lvw_Backup, lvw_Freq);
            for (i = (lvw_Freq.Items.Count - 1); i > 0; i--)
            {
                S = lvw_Freq.Items[i].SubItems[j].Text;
                if ((float.Parse(S) < Min_Val) || (float.Parse(S) > Max_Val))
                {
                    lvw_Freq.Items[i].EnsureVisible();
                    lvw_Freq.Items[i].Remove();
                    lvw_Freq.Refresh();
                    txt_Display.Text = "Processing Assoc.Rule # " + i.ToString();
                    txt_Display.Refresh();
                }
            }
           Colorize_List_View();
           txt_Display.Text = (lvw_Freq.Items.Count).ToString() + " Association Rules Generated"; 
        }

        private bool Is_Subset(String Msg, String Find)
        {
            bool Ret = true;
            int i;
            for (i = 0; i < Find.Length; i++)
            {
                if (Msg.Contains(Find[i].ToString()) == false)
                {
                    return false;
                }
            }
            return Ret;
        }

        private float Refresh_SAR(int SAR_ID , String Selected_Measure)
        {   
            OleDbCommand CMD_1 = new OleDbCommand();
            OleDbCommand CMD_2 = new OleDbCommand();
            OleDbDataReader R;
            int ID, Total_DB;
            float LSup, RSup, Ret = 0;
            float Support, Confidence, Leverage, Lift, Conviction, All_Confidence;             
            //--------------------------------------------       
            CMD_1.Connection = Con;
            CMD_2.Connection = Con;
            Total_DB = int.Parse(lbl_Total_DB_Records.Text);
            //--------------------------------------------                        
            CMD_1.CommandText = "Select * from SAR "; // +SQL_Ext;
            R = CMD_1.ExecuteReader();
            while(R.Read())
            {
                ID = int.Parse(R[0].ToString ());
                
                LSup = 100 * (float.Parse(Calculate_Frequency(R[1].ToString()).ToString()) / Total_DB); 
                CMD_2.CommandText = "update SAR Set LSup=" + LSup + " Where ID=" + ID;
                CMD_2.ExecuteNonQuery();
                //txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine);
                     //--------------------------------------------      
                RSup = 100 * (float.Parse(Calculate_Frequency(R[2].ToString()).ToString()) / Total_DB); 
                CMD_2.CommandText = "update SAR Set RSup=" + RSup + " Where ID=" + ID;
                CMD_2.ExecuteNonQuery();
                //txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine);
                     //--------------------------------------------      
                Support = 100 * (float.Parse(Calculate_Frequency(R[1].ToString() + R[2].ToString()).ToString()) / Total_DB);
                CMD_2.CommandText = "update SAR Set Support=" + Support + " Where ID=" + ID;
                CMD_2.ExecuteNonQuery();
                //txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine);
                //--------------------------------------------                  
                Confidence = 100 * Support / LSup;
                //if (LSup == 0) Confidence = 100;
                CMD_2.CommandText = "update SAR Set Confidence=" + Confidence + " Where ID=" + ID ;
                CMD_2.ExecuteNonQuery();
                // txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine);
                //--------------------------------------------                  
                Leverage = (Support / 100);
                Leverage =  (Leverage - ((LSup / 100)*(RSup/100)));
                CMD_2.CommandText = "update SAR Set Leverage=" + Leverage + " Where ID=" + ID;
                CMD_2.ExecuteNonQuery();
                // txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine);
                //--------------------------------------------                                  
                Lift = Total_DB * ((Confidence / 100) / (LSup));
                CMD_2.CommandText = "update SAR Set Lift=" + Lift + " Where ID=" + ID;
                CMD_2.ExecuteNonQuery();
                // txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine);
                //--------------------------------------------                                                  
                Conviction = ((1 - (RSup/100)) / (1-(Confidence/100)));
                CMD_2.CommandText = "update SAR Set Conviction=" + Conviction + " Where ID=" + ID;
                CMD_2.ExecuteNonQuery();
                // txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine);
                //--------------------------------------------                                                                
                All_Confidence = Measure_All_Confidence(R[1].ToString() + R[2].ToString());
                CMD_2.CommandText = "update SAR Set All_Confidence=" + All_Confidence + " Where ID=" + ID;
                CMD_2.ExecuteNonQuery();
                // txt_Display.AppendText(CMD_2.CommandText + Environment.NewLine);
                //--------------------------------------------                  
                if (SAR_ID == ID)
                {
                    switch (Selected_Measure)
                    {
                        case "Confidence":
                            {
                                Ret = Confidence;
                                break;
                            }
                        case "Leverage":
                            {
                                Ret = Leverage; 
                                break;
                            }
                        case "Lift":
                            {
                                Ret = Lift;
                                break;
                            }
                        case "Conviction":
                            {
                                Ret = Conviction;
                                break;
                            }
                        case "All-Confidence":
                            {
                                Ret = All_Confidence;
                                break;
                            }
                    }                    
                }
                //--------------------------------------------                  
                //txt_Display.AppendText(Environment.NewLine + "SAR ID:" + ID + " LSup:" + LSup + " RSup:" + RSup + " Sup:" + Support + " Conf:" + Confidence + " Leverg:" + Leverage + " Lift: " + Lift + " Conv:" + Conviction + " All-Conf:" + All_Confidence);                                
            }
            R.Close();
            txt_Display.AppendText(Environment.NewLine);
            return Ret;
        }
                
        private char[] Sort_SAR_Items(String SortOrder)
        {
            OleDbCommand CMD_1 = new OleDbCommand();
            OleDbDataReader R;
            String S = "";
            String Ret = "";
            bool F = false;
            int i, j;
            //--------------------------------------------                                                                        
            CMD_1.CommandText = "Select * from SAR";
            CMD_1.Connection = Con;
            R = CMD_1.ExecuteReader();
            while (R.Read())
            {
                S += R[1].ToString() + R[2].ToString();
            }
            R.Close();
            // construction of unique characters string. (read duplicate character only once)
            for (i = 0; i < S.Length; i++)
            {
                F = false;
                for (j = S.Length - 1; j > i; j--)
                {
                    if (S[i] == S[j])
                    {
                        F = true;
                        continue;
                    }
                }
                if (F == false)
                    Ret += S[i];
            }
            //////////////////////
            // sort it now along with its frequencey
            int[] Freq = new int[Ret.Length];
            int tmp; char t;
            for (i = 0; i < Ret.Length; i++)
            {
                Freq[i] = Calculate_Frequency(Ret[i].ToString());
            }

            char[] Ar = Ret.ToCharArray();

            for (i = 0; i < Freq.Length - 1; i++)
            {
                for (j = i + 1; j < Freq.Length; j++)
                    if (Freq[i] < Freq[j])
                    {
                        tmp = Freq[i];
                        t = Ar[i];
                        Freq[i] = Freq[j];
                        Ar[i] = Ar[j];
                        Freq[j] = tmp;
                        Ar[j] = t;
                    }
            }
            /*
            for (i = 0; i < Freq.Length; i++)
                txt_Display.AppendText(Ar[i].ToString() + " = " + Freq[i].ToString() + "   ");
            txt_Display.AppendText(Environment.NewLine + "----------------------" + Environment.NewLine);
            */
            // by default SortOrder above was Desc                        
            if (SortOrder.Contains("Asc"))
            {
                char[] ArRev = new char[Ar.Length];
                for (i = 0; i < Freq.Length; i++)
                {
                    ArRev[i] = Ar[Freq.Length - i - 1];
                }
                return ArRev; // in Asc order 
            }

            return Ar; // in Desc order
        }

        private void chk_FOC_SAR_Asc_Desc_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_FOC_SAR_Asc_Desc.Text == "FOC SAR Asc")
                chk_FOC_SAR_Asc_Desc.Text = "FOC SAR Desc";
            else
                chk_FOC_SAR_Asc_Desc.Text = "FOC SAR Asc";
        }

        private String[] Extract_Character_No(char[] Ar, String Antec, String Conseq)
        {
            int tmp = -1; String[] Ret = new String[2];
            int i, j;

            for (i = 0; i < Ar.Length; i++)
            {
                for (j = 0; j < Antec.Length; j++)
                {
                    if (Antec[j] == Ar[i])
                    {
                        tmp = i;
                        Ret[0] = "Antecedent";
                        Ret[1] = j.ToString();
                        i = Ar.Length + 1;
                        j = Antec.Length + 1;
                    }
                }
            }
            //////////////////////////////////////////
            for (i = 0; i < Ar.Length; i++)
            {
                for (j = 0; j < Conseq.Length; j++)
                {
                    if ((Conseq[j] == Ar[i]) && (i < tmp))
                    {
                        tmp = i;
                        Ret[0] = "Consequent";
                        Ret[1] = j.ToString();
                        i = Ar.Length + 1;
                        j = Conseq.Length + 1;
                    }
                }
            }
            //////////////////////////////////////////
            return Ret;
        }

        private void Generate_XL(ListView LVW)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;           

            int Cl, Rw;
            String Start_Range = "A2";
            String End_Range = String.Empty;
            Char ch = Start_Range[0];
            for (Cl = 0; Cl < LVW.Columns.Count; Cl++)
                ch++;
            End_Range = ch.ToString() + "1";

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();                
                oXL.Visible = true;

                //Get a new workbook.            
                //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oWB = (Excel._Workbook)(oXL.Workbooks.Open(XL_File, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0));
                oWB.Sheets.Add(Missing.Value, Missing.Value , 1,Excel.XlSheetType.xlWorksheet);


                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                //oSheet.Name = "Assoc.Rules.Result";

                oSheet.Cells[1, 1] = "FOC"; // Frequency of Count 
                oSheet.Cells[1, 2] = lst_Support.Text;
                oSheet.Cells[1, 4] = "|DB Transactions|";  // Total DB Transactions
                oSheet.Cells[1, 5] = lbl_Total_DB_Records.Text; 
                oSheet.Cells[1, 6] = "|DB Items|"; // Total DB Items
                oSheet.Cells[1, 7] = lbl_Total_Items.Text; 
                oSheet.Cells[1, 8] = "SAR FOC Order"; // SAR sorted by FOC while in pruning  
                if (chk_FOC_SAR_Asc_Desc.Checked == true) 
                  oSheet.Cells[1, 9] = "Desc"; 
                else
                    oSheet.Cells[1, 9] = "Asc"; 
                //Add table headers going cell by cell.
                for (Cl = 0; Cl < LVW.Columns.Count; Cl++)
                {
                    oSheet.Cells[2, Cl + 1] = LVW.Columns[Cl].Text; // Columns Name                    
                }

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range(Start_Range, End_Range).Font.Bold = true;
                oSheet.get_Range(Start_Range, End_Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                for (Rw = 0; Rw < LVW.Items.Count; Rw++)
                {
                    for (Cl = 0; Cl < LVW.Columns.Count; Cl++)
                    {
                        oSheet.Cells[Rw + 3, Cl + 1] = LVW.Items[Rw].SubItems[Cl].Text;
                    }
                }

                //AutoFit columns A:D.
                oRng = oSheet.get_Range(Start_Range, End_Range);
                oRng.EntireColumn.AutoFit();

                //Manipulate a variable number of columns for Quarterly Sales Data.
                //DisplayQuarterlySales(oSheet);

                //Make sure Excel is visible and give the user control 
                //of Microsoft Excel's lifetime.
                oXL.Visible = true;
                oXL.UserControl = false;

                oWB.SaveAs(XL_File,Type.Missing, Type.Missing, Type.Missing,Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing);
                oWB.Close(true, XL_File, Type.Missing);
                oXL.Workbooks.Close();                
                oXL.Quit(); 
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");
            }            
        }

        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {
            Excel._Workbook oWB;
            //Excel.Series oSeries;
            Excel.Range oResizeRange;
            Excel._Chart oChart;            
            //int iNumQtrs;

            //Determine how many quarters to display data for.
            //for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            //{            
            //    DialogResult iRet = DialogResult.Yes;
            //    if (iRet == DialogResult.Yes) break;
            //}
                        
            //Starting at E1, fill headers for the number of columns selected.
            //oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            //oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            //Change the Orientation and WrapText properties for the headers.
            //oResizeRange.Orientation = 38;
            //oResizeRange.WrapText = true;

            //Fill the interior color of the headers.
            //oResizeRange.Interior.ColorIndex = 36;

           //Fill the columns with a formula and apply a number format.
           // oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
           // oResizeRange.Formula = "=RAND()*100";
           // oResizeRange.NumberFormat = "$0.00";

            //Apply borders to the Sales data and headers.
            oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value,4);
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            //Add a Totals formula for the sales data and apply a border.
           // oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
           // oResizeRange.Formula = "=SUM(E2:E6)";
           // oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble;
           // oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThick;

            //Add a Chart for the selected data.
            oWB = (Excel._Workbook)oWS.Parent;
            oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,Missing.Value, Missing.Value);
            
            //Use the ChartWizard to create a new chart from the selected data.
            oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(Missing.Value, 4);
            oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            /*
            oSeries = (Excel.Series)oChart.SeriesCollection(1);
            oSeries.XValues = oWS.get_Range("A2", "A6");
            for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            {
                oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
                String seriesName;
                seriesName = "=\"Q";
                seriesName = String.Concat(seriesName, iRet);
                seriesName = String.Concat(seriesName, "\"");
                oSeries.Name = seriesName;
            }
            */
            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            //Move the chart so as not to cover your data.
            oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
            oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
            oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
            
        }

        private void mnu_Excel_Click(object sender, System.EventArgs e)
        {
            if (rad_Result_XL.Checked == true)
                Generate_XL(lvw_Report);
            else if(rad_FI_AR_XL.Checked == true)
                Generate_XL(lvw_Freq);            
        }

        private void Fill_Progress_Bar()
        {
            int i, UV;
            float Perc;
            UV = lbx_Measure.Items.Count * lbx_Weights.Items.Count; ;
            i = lvw_Report.Items.Count;
            pbr_Busy.Value = 0;
            pbr_Busy.Maximum = UV;
            Perc = float.Parse(i.ToString()) / UV;
            Perc = Perc * 100;
            if (Perc >= 50) lbl_Progress.BackColor = pbr_Busy.ForeColor;
            lbl_Progress.Text = Perc.ToString() + "%";
            pbr_Busy.Value = i;
            grp_Progress.Refresh();
        }

        private void chk_Populate_Wt_CheckedChanged(object sender, EventArgs e)
        {
            lbx_Weights.Items.Clear();
            lbx_Weights.Items.Add("SumOf");
            if (chk_Populate_Wt.Checked == true)
            {
                lbx_Weights.Items.Add("Mean");
                lbx_Weights.Items.Add("Mode");
                lbx_Weights.Items.Add("Median");
            }
        }




        private void Generate_Tab_Text(ListView LVW)
        {
            int Col_Count = LVW.Columns.Count;
            int Cl, Rw;
            int tmpC = LVW.Items.Count;
            String S = String.Empty;
            StreamWriter SW = File.AppendText(Application.StartupPath.ToString() + "\\" + "naeem.tab");

            for (Cl = 0; Cl < Col_Count; Cl++)
            {
                S = S + LVW.Columns[Cl].Text + ","; // Columns Name                    
            }
            SW.WriteLine(S);
            S = "";

            for (Rw = 0; Rw < LVW.Items.Count; Rw++)
            {
                for (Cl = 0; Cl < Col_Count; Cl++)
                {
                    S = S + LVW.Items[Rw].SubItems[Cl].Text.ToString() + ",";
                }
                txt_Display.Text = tmpC.ToString();
                tmpC--;
                txt_Display.Refresh();
                SW.WriteLine(S);
                S = "";
            }
            SW.Close();
            txt_Display.Text = "File naeem.tab" + " written..";
        }

    }


}
