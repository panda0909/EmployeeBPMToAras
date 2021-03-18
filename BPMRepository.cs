using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace JPC_Reports.Models.BPM
{
    public class BPMRepository
    {
        private string connectionStr = "";
        public string error_msg = "";
        public BPMRepository(string connect)
        {
            connectionStr = connect;
        }
        
        public DataTable GetEmployeeLevel()
        {
            DataTable dt = RunSql(@"SELECT 
                users_occu.userName as Occu
	                ,users_occu.id as EMP_ID
	                ,users_occu.ldapid as LDAP_ID
                    ,level.levelValue
	                ,level.functionLevelName
	                ,funcDef.functionDefinitionName
	                ,org_unit.organizationUnitName as DeptName
	                ,org_unit_super.organizationUnitName as SuperDeptName
	                ,org_emp.organizationName as OrgName
                    ,users.userName as Manager
                    ,[isMain]
                FROM [Functions] as func
                left join [OrganizationUnit] as org_unit on org_unit.OID = func.organizationUnitOID
                left join [OrganizationUnit] as org_unit_super on org_unit_super.OID = org_unit.superUnitOID
                left join [FunctionDefinition] as funcDef on funcDef.OID = func.definitionOID
                left join Users as users on users.OID = func.specifiedManagerOID
                left join Users as users_occu on users_occu.OID = func.occupantOID
                left join FunctionLevel as level on level.OID = func.approvalLevelOID
                left join Organization as org_emp on org_emp.OID = org_unit.organizationOID
                order by Occu,isMain");

            return dt;
        }
        public DataTable RunSql(string sqlcmd)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionStr))
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                    }

                    SqlCommand cmd = new SqlCommand(sqlcmd, conn);
                    SqlDataAdapter DataAdapter = new SqlDataAdapter();
                    DataAdapter.SelectCommand = cmd;

                    DataSet ds = new DataSet();
                    DataAdapter.Fill(ds);

                    return ds.Tables[0];
                }
            }
            catch (Exception ex)
            {
                error_msg = ex.ToString();
                return null;
            }
        }
    }
}