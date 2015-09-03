using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Linq;
using System.Text;

public partial class _Default : Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        var path = @"C:\Mango.xlsx";
        var connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
        var conn = new OleDbConnection(connStr);
        conn.Open();

        try
        {
            var MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", conn);
            MyCommand.TableMappings.Add("Table", "TestTable");
            var DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);

            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < DtSet.Tables[0].Rows.Count; i++)
            {

                var firstname = DtSet.Tables[0].Rows[i].ItemArray[1];
                var lastname = DtSet.Tables[0].Rows[i].ItemArray[0];
                var jobname = DtSet.Tables[0].Rows[i].ItemArray[2];
                var startdate = DtSet.Tables[0].Rows[i].ItemArray[5];

                sb.Append("select * from employee where lastname = '" + lastname + "' and firstname = '" + firstname + "' and EmployeeHRID =  '' and jobgroupid in (select jobgroupid from jobgroup where name = '" + jobname + "') and hiredate = '" + startdate + "' <br />");


            }

            lbl1.Text = sb.ToString();





        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message.ToString());
        }
        finally
        {
            conn.Close();
        }
    }
}