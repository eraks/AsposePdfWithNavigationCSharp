using System;
using System.Drawing;
using System.Text;
using Aspose.Words;
using System.Data;

namespace pdfNavigationPoc
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            DataSet ds = DataCrearion();

            DataTable dt0 = ds.Tables[0];
            DataTable dt = ds.Tables[1];


            builder.InsertHtml("<br><h2>Society (Public) Review</h2>");
            foreach (DataRow dataRow in dt0.Rows)
            {
                builder.Font.Color = Color.Blue;
                builder.Font.Underline = Underline.Single;
                builder.InsertBreak(BreakType.LineBreak);
                builder.InsertBreak(BreakType.LineBreak);
                builder.InsertHyperlink(dataRow["Ballot"].ToString(), $@"{dataRow["Ballot"].ToString()}"" \o ""Hyperlink Tip", true);

            }

            string previusBallot = "";
            StringBuilder strTable = new StringBuilder();
            foreach (DataRow dataRow in dt.Rows)
            {
                if (previusBallot != dataRow["Ballot"].ToString())
                {
                    if(previusBallot!="")
                    {
                        strTable.Append("</table>");
                        builder.InsertHtml(strTable.ToString());
                        strTable.Clear();
                    }
                    previusBallot = dataRow["Ballot"].ToString();
                    builder.InsertBreak(BreakType.PageBreak);

                    builder.StartBookmark(dataRow["Ballot"].ToString());

                    builder.InsertHtml($"<h3>{dataRow["Ballot"].ToString()} </h3>");
                    builder.EndBookmark(dataRow["Ballot"].ToString());
                    strTable.Append("<br><table border=1 style='padding:4px'><thead><tr><td>Item No.</td><td>Item</td></tr></thead>");
                }
                strTable.Append($"<tr><td>{ dataRow["Number"].ToString() }</td><td>{ dataRow["Description"].ToString() }</td></tr>");
            }

            strTable.Append("</table>");
            builder.InsertHtml(strTable.ToString());
            strTable.Clear();

            doc.Save("OutPut.pdf");
        }

   
        private static DataSet DataCrearion()
        {
            DataSet ds = new DataSet();

            var dt0 = new System.Data.DataTable("LandingTable");

            dt0.Columns.Add("Ballot", typeof(string));
            dt0.Rows.Add(new Object[] { "A06 (21-01)" });
            dt0.Rows.Add(new Object[] { "C09 (21-02)" });

            var dt = new System.Data.DataTable("DetailTable");

            dt.Columns.Add("Number", typeof(int));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Ballot", typeof(string));

            dt.Rows.Add(new Object[] { 1, "Reapproval of A0720/A0720M-2002(2016)E1 Test Method for Ductility of Nonoriented Electrical Steel WK75550", "A06 (21-01)" });
            dt.Rows.Add(new Object[] { 2, "Revision Of A0977/A0977M-2007(2020) Test Method for Magnetic Properties of High-Coercivity Permanent Magnet Materials Using Hysteresigraphs WK75756", "A06 (21-01)" });
            dt.Rows.Add(new Object[] { 3, "Reapproval of A0721/A0721M-2002(2016) Test Method for Ductility of Oriented Electrical Steel WK75551", "A06 (21-01)" });
            dt.Rows.Add(new Object[] { 1, "Specification for Colloidal Silica for Use in Concrete WK60809", "C09 (21-02)" });
            dt.Rows.Add(new Object[] { 2, "Specification for Performance of Supplementary Cementitious Material for Use in Concrete WK70466", "C09 (21-02)" });

            ds.Tables.Add(dt0);
            ds.Tables.Add(dt);

            return ds;
        }
    }
}
