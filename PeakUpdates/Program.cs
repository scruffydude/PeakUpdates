using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Peak_Updater
{
    class Program
    {

        public const string laborPlanLocationation = @"\\chewy.local\bi\BI Community Content\Finance\Labor Models\";
        public static string[] rowlabels = { "Planned Units Shipped (Days)", "Planned Supply Chain Units Ordered (Days)", "Planned Units Shipped (Nights)", "Planned Supply Chain Units Ordered (Nights)" };
        static void Main(string[] args)
        {
            //connect to DEV Server using SQL Server Data Connection
            System.DateTime appstart = System.DateTime.Now;
            //SqlConnection conn = new SqlConnection();

            string updatePath = @"\\CFC1AFS01\Operations-Analytics\Projects\Hourly Peak Update\Peak_hourly_Updates.xlsm";
            string[] warehouses = { "AVP1", "CFC1", "DFW1", "EFC3", "WFC2" };
            


            string ConnectionString = "Data source=WMSSQL-READONLY.chewy.local;"
               + "Initial Catalog=AAD;"
               + "Persist Security Info=True;"
               + "Trusted_connection=true";

            Excel.Application app = new Excel.Application();
            Excel.Workbook Peakwkbk = app.Workbooks.Open(updatePath);
            Excel.Worksheet Datawksht = Peakwkbk.Sheets.Item["DataUpdate"];
            Excel.Worksheet LPInfo = Peakwkbk.Sheets.Item["21DPInfo"];
            app.Calculation = Excel.XlCalculation.xlCalculationManual;
            app.Visible = false;
            app.DisplayAlerts = false;

            Excel.Workbook LaborPlan = null;

            string[] laborplans = Directory.GetFiles(laborPlanLocationation);

            int lastRow = Datawksht.UsedRange.Rows.Count;
            Excel.Range r1 = Datawksht.Cells[2, 1];
            Excel.Range r2 = Datawksht.Cells[lastRow, 5];
            Datawksht.Range[r1, r2].Value = "";

            foreach(string lp in laborplans)
            {
                int destCol = 2;
                if (System.IO.Path.GetFileName(lp).Contains("Labor Model"))
                {
                    try
                    {
                        LaborPlan = app.Workbooks.Open(lp, false, true);
                        
                    }
                    catch
                    {
                        continue;
                    }

                    Console.WriteLine("Gathering Info from {0}", lp);
                    Excel.Worksheet OBLaborPlan = LaborPlan.Sheets.Item["OB Daily Plan"];

                    switch (System.IO.Path.GetFileName(lp))
                    {
                        case "AVP1 Labor Model 2017.xlsx":
                            destCol = 2;
                            break;
                        case "CFC1 Labor Model 2017.xlsx":
                            destCol = 3;
                            break;
                        case "DFW1 Labor Model 2017.xlsx":
                            destCol = 5;
                            break;
                        case "EFC3 Labor Model 2017.xlsx":
                            destCol = 6;
                            break;
                        case "WFC2 Labor Model 2017.xlsx":
                            destCol = 7;
                            break;
                        default:
                            break;
                    }

                    grabLaborPlanData(OBLaborPlan, LPInfo, destCol);

                    OBLaborPlan = null;
                    LaborPlan.Close(false);
                    LaborPlan = null;
                }
            }


            GatherSQLData(ConnectionString, Datawksht);

            if(System.DateTime.Now.Hour == 5)
            {
                LPInfo.Cells[7, 2].value = System.DateTime.Today;
                UpdateBLInfo(ConnectionString, LPInfo);
            }

            foreach(string wh in warehouses)
            {
                Console.WriteLine("Sending Mail for {0}", wh);
                app.Run("MailUpdate", wh);
            }
            
                

            //Console.ReadLine();

            app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Datawksht = null;
            Peakwkbk.Close(true);
            app.Quit();
            Peakwkbk = null;
            app = null;
        }

        public static void grabLaborPlanData(Excel.Worksheet OBLaborPlan, Excel.Worksheet LPInfo, int destCol, int destRow = 2)
        {
            foreach (string label in rowlabels)
            {
                for (int x = 1; x < OBLaborPlan.UsedRange.Rows.Count; x++)
                {
                    if (OBLaborPlan.Cells[x, 2].value == label)
                    {
                        LPInfo.Cells[destRow, destCol].value = Math.Round(OBLaborPlan.Cells[x, 2 + DateTime.Now.DayOfYear].value);
                        destRow++;
                    }
                }
            }
        }
        public static void GatherSQLData(string connectionString, Excel.Worksheet Datawksht)
        {
            
            string queryString = "Use AAD declare @start datetime = dateadd(hour, -24, getdate()),@end datetime = getdate() + 1 " +

                            "select cast(isnull(a.review_date, b.review_date) as date) review_date, isnull(a.review_hour, b.review_hour) review_hour, isnull(a.wh_id, b.wh_id) wh_id, isnull(a.ordered_units, 0) ordered_units, isnull(b.shipped_units, 0) shipped_units " +

                            "from  (select cast(o.arrive_date as date) review_date, datepart(hour, o.arrive_date) review_hour, o.wh_id, sum(d.qty) ordered_units " +

                            "from AAD.dbo.t_order o left join AAD.dbo.t_order_detail d on d.wh_id = o.wh_id and o.order_number = d.order_number where o.arrive_date >= @start and o.arrive_date < @end and o.type_id = 31 " +
                            "group by cast(o.arrive_date as date), datepart(hour, o.arrive_date), o.wh_id ) a " +

                            "full outer join (select cast(t.start_tran_date_time as date) review_date, datepart(hour, t.start_tran_date_time) review_hour, t.wh_id, sum(t.tran_qty) shipped_units " +
                            "from AAD.dbo.t_tran_log t where t.start_tran_date_time >= @start and t.tran_type = '341' " +

                            "group by cast(t.start_tran_date_time as date), datepart(hour, t.start_tran_date_time), t.wh_id ) b on a.review_date = b.review_date and a.wh_id = b.wh_id and a.review_hour = b.review_hour order by a.review_date, a.review_hour, a.wh_id";

            using (SqlConnection connection =
                   new SqlConnection(connectionString))
            {
                SqlCommand command =
                    new SqlCommand(queryString, connection);
                connection.Open();

                SqlDataReader reader = command.ExecuteReader();

                Console.WriteLine("{0}, {1}, {2}, {3}, {4}", reader.GetName(0), reader.GetName(1), reader.GetName(2), reader.GetName(3), reader.GetName(4));
                for(int x = 0; x<5;x++)
                {
                    Datawksht.Cells[1, x+1].value = reader.GetName(x);
                }
                int y = 2;
                // Call Read before accessing data.
                while (reader.Read())
                {
                    ReadEachRow((IDataRecord)reader, Datawksht, y);
                    y++;
                }

                // Call Close when done reading.
                reader.Close();
            }

            
        }

        public static void ReadEachRow(IDataRecord record, Excel.Worksheet data, int y)
        {
            for(int i = 1; i<6; i++)
            {
                data.Cells[y, i].value = record[i-1];
            }
            Console.WriteLine("{0}, {1}, {2}, {3}, {4}", record[0], record[1], record[2], record[3], record[4]);
        }

        public static void UpdateBLInfo(string connectionString, Excel.Worksheet LPInfo)
        {
            string queryString = "USE AAD SELECT c.wh_id,(GETDATE()) date_pulled,SUM(total_units) total_units " +
                "FROM (SELECT pkc.wh_id, (CASE WHEN flw.zone IS NOT NULL AND COUNT(DISTINCT pkd.item_number) = 1 THEN 'Flow Singles' WHEN flw.zone IS NOT NULL THEN 'Flow Multis' WHEN cpb.print_timing = 'POSTPRINT' THEN 'Singles' ELSE 'Multis' END) AS group_name, SUM(pkd.planned_quantity)AS total_units, pkc.container_id AS total_containers, SUM((CASE WHEN(ord.arrive_date < ord.cut_datetime AND ord.pull_datetime < GETDATE() " +
                "OR ord.arrive_date < (ord.cut_datetime - 1)) THEN pkd.planned_quantity ELSE 0 END)) AS aged_units, COUNT(DISTINCT CASE WHEN(ord.arrive_date < ord.cut_datetime AND ord.pull_datetime < GETDATE() OR ord.arrive_date < (ord.cut_datetime - 1)) THEN pkc.container_id ELSE NULL END) AS aged_containers " +
                "FROM v_order_cutoff ord(NOLOCK) INNER JOIN v_pick_container_active pkc (NOLOCK)ON pkc.order_number = ord.order_number AND pkc.wh_id = ord.wh_id INNER JOIN v_pick_detail_active pkd(NOLOCK) ON pkd.container_id = pkc.container_id AND pkd.wh_id = pkc.wh_id LEFT JOIN t_container_print_detail cpd(nolock) ON pkd.container_id = cpd.container_id " +
                "LEFT JOIN t_container_print cp(nolock) ON cp.batch_id = cpd.batch_id AND cp.wh_id = ord.wh_id LEFT JOIN t_container_print_batch_type cpb(nolock) ON cp.container_print_batch_type_id = cpb.unique_id OUTER APPLY (SELECT TOP 1 z.zone, p.status, p.wh_id FROM v_pick_detail_active p(NOLOCK) INNER JOIN t_zone_loca zlc(NOLOCK) ON zlc.location_id = p.pick_location AND zlc.wh_id = p.wh_id " +
                "INNER JOIN t_zone z(NOLOCK) ON z.zone = zlc.zone AND z.wh_id = zlc.wh_id AND z.zone_type = 'WALL' WHERE p.container_id = pkc.container_id AND p.wh_id = pkc.wh_id ORDER BY(CASE WHEN p.status = 'RELEASED' THEN 0 ELSE 1 END), z.sequence) flw WHERE ord.status <> 'S' GROUP BY pkc.wh_id, flw.zone, cpb.print_timing, pkc.container_id) c GROUP BY c.wh_id";
            using (SqlConnection connection =
                               new SqlConnection(connectionString))
            {
                SqlCommand command =
                    new SqlCommand(queryString, connection);
                connection.Open();

                SqlDataReader reader = command.ExecuteReader();

                //Console.WriteLine("{0}, {1}, {2}, {3}, {4}", reader.GetName(0), reader.GetName(1), reader.GetName(2), reader.GetName(3), reader.GetName(4));

                int y = 2;
                // Call Read before accessing data.
                while (reader.Read())
                {
                    setBacklogs((IDataRecord)reader, LPInfo, y);
                    y++;
                }

                // Call Close when done reading.
                reader.Close();
            }
        }
        public static void setBacklogs( IDataRecord record, Excel.Worksheet LPInfo, int y)
        {

                LPInfo.Cells[6,y].value = record[2];
        }
    }  
}
