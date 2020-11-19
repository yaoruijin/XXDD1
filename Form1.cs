using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace XXDD
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void button5_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		//public string constr = "server=YRJ;database=UFDATA_003_2016;uid=sa;pwd=111";
		public string constr = "server=FWQ;database=UFDATA_003_2016;uid=yq;pwd=aa*123";
		private void button1_Click(object sender, EventArgs e)
		{
			comboBox1.Items.Clear();
			dataGridView1.Rows.Clear();
			dataGridView2.Rows.Clear();


			SqlConnection sqlc = new SqlConnection(constr);

			sqlc.Open();
			string findstr1 = " SELECT  [cCode] FROM [dbo].[rdrecord09] WHERE  [cSource]='库存' AND  [dDate]='" + dateTimePicker1.Value.ToString("yyyy-MM-dd 00:00:00") + "' ";
			SqlCommand com1 = new SqlCommand(findstr1, sqlc);
			SqlDataAdapter sqlda1 = new SqlDataAdapter(com1);

			DataSet ds1 = new DataSet();
			sqlda1.Fill(ds1, "danhao");
			sqlc.Close();

			for (int x = 0; x < ds1.Tables["danhao"].Rows.Count; x++)
			{
				comboBox1.Items.Add(ds1.Tables["danhao"].Rows[x][0]);
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{


			SqlConnection sqlc2 = new SqlConnection(constr);

			sqlc2.Open();
			string findstr2 = " SELECT [RdRecord09].[ID],[dDate],[cCode],[cRdCode],rdrecords09.[cInvCode],Inventory.[cInvName],Inventory.[cInvStd],[cWhCode],[cDefine11], rdrecords09.cDefine28,rdrecords09.iQuantity,rdrecords09.iUnitCost,rdrecords09.iPrice FROM[RdRecord09],[rdrecords09],[Inventory]  where[RdRecord09].[ID] = rdRecords09.[ID] and rdrecords09.cInvCode = Inventory.cInvCode and cCode = '" + comboBox1.Text + "' ";
			SqlCommand com2 = new SqlCommand(findstr2, sqlc2);
			SqlDataAdapter sqlda2 = new SqlDataAdapter(com2);

			DataSet ds2 = new DataSet();
			sqlda2.Fill(ds2, "cangku");
			sqlc2.Close();

			for (int x = 0; x < ds2.Tables["cangku"].Rows.Count; x++)
			{
				dataGridView1.Rows.Add();
				dataGridView1.Rows[x].Cells[0].Value = ds2.Tables["cangku"].Rows[x][4].ToString();
				dataGridView1.Rows[x].Cells[1].Value = ds2.Tables["cangku"].Rows[x][5].ToString();
				dataGridView1.Rows[x].Cells[2].Value = ds2.Tables["cangku"].Rows[x][6].ToString();
				dataGridView1.Rows[x].Cells[3].Value = ds2.Tables["cangku"].Rows[x][7].ToString();
				dataGridView1.Rows[x].Cells[4].Value = ds2.Tables["cangku"].Rows[x][8].ToString();
				dataGridView1.Rows[x].Cells[5].Value = ds2.Tables["cangku"].Rows[x][9].ToString();
				dataGridView1.Rows[x].Cells[6].Value = Math.Round(Convert.ToDouble(ds2.Tables["cangku"].Rows[x][10].ToString()), 4);

				dataGridView1.Rows[x].Cells[7].Value = Math.Round(Convert.ToDouble(ds2.Tables["cangku"].Rows[x][11].ToString()), 5);
				dataGridView1.Rows[x].Cells[8].Value = Math.Round(Convert.ToDouble(ds2.Tables["cangku"].Rows[x][12].ToString()), 2);
			}



		}

		private void textBox2_TextChanged(object sender, EventArgs e)
		{
			if (textBox2.Text != null)
			{



				SqlConnection sqlc3 = new SqlConnection(constr);

				sqlc3.Open();
				string findstr3 = " SELECT  [cCusCode],[cCusName],[cCusAbbName], cCusCreditCompany  FROM [Customer] where CHARINDEX('" + textBox2.Text + "',cCusName)> 0 ";
				SqlCommand com3 = new SqlCommand(findstr3, sqlc3);
				SqlDataAdapter sqlda3 = new SqlDataAdapter(com3);

				DataSet ds3 = new DataSet();
				sqlda3.Fill(ds3, "kehu");
				sqlc3.Close();


				for (int x = 0; x < ds3.Tables["kehu"].Rows.Count; x++)
				{
					dataGridView3.Rows.Add();
					dataGridView3.Rows[x].Cells[0].Value = ds3.Tables["kehu"].Rows[x][0].ToString();
					dataGridView3.Rows[x].Cells[1].Value = ds3.Tables["kehu"].Rows[x][1].ToString();
					dataGridView3.Rows[x].Cells[2].Value = ds3.Tables["kehu"].Rows[x][2].ToString();
					dataGridView3.Rows[x].Cells[3].Value = ds3.Tables["kehu"].Rows[x][3].ToString();

				}

				for (int x = 0; x < ds3.Tables["kehu"].Rows.Count; x++)
				{
					for (int x1 = 0; x1 < ds3.Tables["kehu"].Rows.Count; x1++)

						if (dataGridView3.Rows[x].Cells[3].Value.ToString() == ds3.Tables["kehu"].Rows[x1][0].ToString())
						{
							dataGridView3.Rows[x].Cells[4].Value = ds3.Tables["kehu"].Rows[x1][2].ToString();
						}
				}



			}
		}

		private void textBox3_TextChanged(object sender, EventArgs e)
		{

			if (textBox3.Text != null)
			{



				SqlConnection sqlc4 = new SqlConnection(constr);

				sqlc4.Open();
				string findstr4 = " SELECT [citemcode],[citemname],[主管业务员] FROM [fitemss00]  where CHARINDEX('" + textBox3.Text + "',citemname)> 0 ";
				SqlCommand com4 = new SqlCommand(findstr4, sqlc4);
				SqlDataAdapter sqlda4 = new SqlDataAdapter(com4);

				DataSet ds4 = new DataSet();
				sqlda4.Fill(ds4, "项目");
				sqlc4.Close();


				for (int x = 0; x < ds4.Tables["项目"].Rows.Count; x++)
				{
					dataGridView4.Rows.Add();
					dataGridView4.Rows[x].Cells[0].Value = ds4.Tables["项目"].Rows[x][0].ToString();
					dataGridView4.Rows[x].Cells[1].Value = ds4.Tables["项目"].Rows[x][1].ToString();
					dataGridView4.Rows[x].Cells[2].Value = ds4.Tables["项目"].Rows[x][2].ToString();


				}



			}

		}

		private void button6_Click(object sender, EventArgs e)
		{
			if (dataGridView3.SelectedRows.Count == 1)
			{
				int introw = dataGridView3.SelectedCells[0].RowIndex;

				label11.Text = dataGridView3.CurrentRow.Cells[0].Value.ToString();
				label12.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();

				label13.Text = dataGridView3.CurrentRow.Cells[3].Value.ToString();
				label14.Text = dataGridView3.CurrentRow.Cells[4].Value.ToString();

			}
		}

		private void button7_Click(object sender, EventArgs e)
		{
			if (dataGridView4.SelectedRows.Count == 1)
			{
				int introw = dataGridView4.SelectedCells[0].RowIndex;

				label15.Text = dataGridView4.CurrentRow.Cells[0].Value.ToString();
				label16.Text = dataGridView4.CurrentRow.Cells[1].Value.ToString();



			}
		}

		private void button3_Click(object sender, EventArgs e)
		{

			
			File.Delete(@"E:\\销售订单\\导出.xls");

			string lujing = System.Environment.CurrentDirectory;

			if (File.Exists(@"" + lujing + "\\销售订单.xls"))
			{


				Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

				Microsoft.Office.Interop.Excel.Workbook workbook = xlApp.Workbooks.Open(@"" + lujing + "\\销售订单.xls");



				Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1





				int r = 0;

				for (r = 1; r < dataGridView1.Rows.Count - 1; r++)
				{


					worksheet.Cells[r + 1, 1] = "YQ" + DateTime.Now.ToString("yyyyMMdd") + comboBox1.Text.ToString().Substring(5, 5);
					worksheet.Cells[r + 1, 1].NumberFormat = "@";


					worksheet.Cells[r + 1, 2] = DateTime.Now.ToString("yyyy-MM-dd");        //日期
																							//  worksheet.Cells[r + 1, 2].NumberFormat = "@";

					worksheet.Cells[r + 1, 3] = "普通销售";
					worksheet.Cells[r + 1, 3].NumberFormat = "@";

					worksheet.Cells[r + 1, 4] = "01";
					worksheet.Cells[r + 1, 4].NumberFormat = "@";

					worksheet.Cells[r + 1, 5] = label11.Text;
					worksheet.Cells[r + 1, 5].NumberFormat = "@";

					worksheet.Cells[r + 1, 6] = label12.Text;
					worksheet.Cells[r + 1, 6].NumberFormat = "@";


					worksheet.Cells[r + 1, 7] = "02";
					worksheet.Cells[r + 1, 7].NumberFormat = "@";

					worksheet.Cells[r + 1, 13] = "人民币";
					worksheet.Cells[r + 1, 13].NumberFormat = "@";

					worksheet.Cells[r + 1, 14] = "1";
					worksheet.Cells[r + 1, 14].NumberFormat = "@";


					worksheet.Cells[r + 1, 15] = "13";
					worksheet.Cells[r + 1, 15].NumberFormat = "@";

					worksheet.Cells[r + 1, 23] = textBox7.Text;
					worksheet.Cells[r + 1, 23].NumberFormat = "@";

					worksheet.Cells[r + 1, 50] = "1";
					worksheet.Cells[r + 1, 50].NumberFormat = "@";



					worksheet.Cells[r + 1, 21] = textBox7.Text;
					worksheet.Cells[r + 1, 21].NumberFormat = "@";

					worksheet.Cells[r + 1, 25] = comboBox1.Text;
					worksheet.Cells[r + 1, 25].NumberFormat = "@";

					worksheet.Cells[r + 1, 26] = label16.Text;
					worksheet.Cells[r + 1, 26].NumberFormat = "@";//项目名称


					// worksheet.Cells[r + 1, 34] = textBox16.Text;
					// worksheet.Cells[r + 1, 34].NumberFormat = "@";//支付比例



					//  worksheet.Cells[r + 1, 35] = textBox4.Text;
					// worksheet.Cells[r + 1, 35].NumberFormat = "@";

					// worksheet.Cells[r + 1, 36] = worksheet.Cells[r + 1, 1];
					// worksheet.Cells[r + 1, 36].NumberFormat = "@";


					worksheet.Cells[r + 1, 43] = DateTime.Now.ToString("yyyy-MM-dd");
					worksheet.Cells[r + 1, 43].NumberFormat = "@";

					worksheet.Cells[r + 1, 44] = DateTime.Now.ToString("yyyy-MM-dd");
					worksheet.Cells[r + 1, 44].NumberFormat = "@";


					worksheet.Cells[r + 1, 45] = dataGridView1.Rows[r - 1].Cells[0].Value;
					worksheet.Cells[r + 1, 45].NumberFormat = "@";

					worksheet.Cells[r + 1, 46] = dataGridView1.Rows[r - 1].Cells[2].Value;
					worksheet.Cells[r + 1, 46].NumberFormat = "@";

					worksheet.Cells[r + 1, 47] = dataGridView1.Rows[r - 1].Cells[1].Value;
					worksheet.Cells[r + 1, 47].NumberFormat = "@";

					worksheet.Cells[r + 1, 48] = dataGridView1.Rows[r - 1].Cells[6].Value;
					worksheet.Cells[r + 1, 48].NumberFormat = "@";

					worksheet.Cells[r + 1, 61] = (double)dataGridView1.Rows[r - 1].Cells[8].Value / 1.13;
					worksheet.Cells[r + 1, 61].NumberFormat = "@";

					worksheet.Cells[r + 1, 55] = (double)dataGridView1.Rows[r - 1].Cells[8].Value / 1.13;
					worksheet.Cells[r + 1, 55].NumberFormat = "@";


					worksheet.Cells[r + 1, 54] = (double)dataGridView1.Rows[r - 1].Cells[8].Value / (double)dataGridView1.Rows[r - 1].Cells[6].Value;
					worksheet.Cells[r + 1, 54].NumberFormat = "@";

					worksheet.Cells[r + 1, 62] = (double)dataGridView1.Rows[r - 1].Cells[8].Value - (double)dataGridView1.Rows[r - 1].Cells[8].Value / 1.13;
					worksheet.Cells[r + 1, 62].NumberFormat = "@";

					worksheet.Cells[r + 1, 56] = (double)dataGridView1.Rows[r - 1].Cells[8].Value - (double)dataGridView1.Rows[r - 1].Cells[8].Value / 1.13;
					worksheet.Cells[r + 1, 56].NumberFormat = "@";

					worksheet.Cells[r + 1, 60] = (double)dataGridView1.Rows[r - 1].Cells[8].Value / (double)dataGridView1.Rows[r - 1].Cells[6].Value / 1.13;
					worksheet.Cells[r + 1, 60].NumberFormat = "@";

					worksheet.Cells[r + 1, 59] = (double)dataGridView1.Rows[r - 1].Cells[8].Value / (double)dataGridView1.Rows[r - 1].Cells[6].Value / 1.13;
					worksheet.Cells[r + 1, 59].NumberFormat = "@";


					worksheet.Cells[r + 1, 63] = (double)dataGridView1.Rows[r - 1].Cells[8].Value;
					worksheet.Cells[r + 1, 63].NumberFormat = "@";

					worksheet.Cells[r + 1, 57] = (double)dataGridView1.Rows[r - 1].Cells[8].Value;
					worksheet.Cells[r + 1, 57].NumberFormat = "@";

					worksheet.Cells[r + 1, 71] = "13";
					worksheet.Cells[r + 1, 71].NumberFormat = "@";

					worksheet.Cells[r + 1, 89] = "00";
					worksheet.Cells[r + 1, 89].NumberFormat = "@";

					worksheet.Cells[r + 1, 90] = label16.Text;
					worksheet.Cells[r + 1, 90].NumberFormat = "@";

					worksheet.Cells[r + 1, 91] = "招标项目";
					worksheet.Cells[r + 1, 91].NumberFormat = "@";

					worksheet.Cells[r + 1, 88] = label15.Text;
					worksheet.Cells[r + 1, 88].NumberFormat = "@";


					// worksheet.Cells[r + 1, 101] = dataGridView1.Rows[r - 1].Cells[15].Value;
					// worksheet.Cells[r + 1, 101].NumberFormat = "@";

					worksheet.Cells[r + 1, 78] = textBox7.Text;
					worksheet.Cells[r + 1, 78].NumberFormat = "@";   //项目名称

					worksheet.Cells[r + 1, 79] = comboBox1.Text;
					worksheet.Cells[r + 1, 79].NumberFormat = "@"; //供货单号

					//  worksheet.Cells[r + 1, 80] = dataGridView1.Rows[r - 1].Cells[1].Value;
					//  worksheet.Cells[r + 1, 80].NumberFormat = "@";   //项目单位


					worksheet.Cells[r + 1, 104] = DateTime.Now.ToString("yyyy-MM-dd");
					worksheet.Cells[r + 1, 104].NumberFormat = "@";

					worksheet.Cells[r + 1, 103] = DateTime.Now.ToString("yyyy-MM-dd");
					worksheet.Cells[r + 1, 103].NumberFormat = "@";




				}




				try
				{

					workbook.Saved = true;
					workbook.SaveAs(@"E:\销售订单\导出.xls");

				}
				catch
				{
				}

				finally
				{
					workbook.Close();
					workbook = null;
					//关闭EXCEL的提示框
					xlApp.DisplayAlerts = false;
					//Excel从内存中退出
					xlApp.Quit();

					xlApp = null;

					GC.Collect();//强行销毁

				}



			}
		
		MessageBox.Show("文件已输出完毕！");
		}
	}
}