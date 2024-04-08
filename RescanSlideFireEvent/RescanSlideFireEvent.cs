using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LSSERVICEPROVIDERLib;
using LSExtensionWindowLib;
using Oracle.ManagedDataAccess.Client;
using Patholab_DAL_V1;
using Patholab_Common;
using System.Runtime.InteropServices;

namespace RescanSlideFireEvent
{
    [ComVisible(true)]
    [ProgId("RescanSlideFireEvent.RescanSlideFireEvent")]
    public partial class RescanSlideFireEvent : UserControl, IExtensionWindow
    {
        public RescanSlideFireEvent()
        {
            InitializeComponent();
            aliquots = new List<ALIQUOT>();
        }

        #region Private Members

        private INautilusUser _ntlsUser;
        private IExtensionWindowSite2 _ntlsSite;

        private INautilusProcessXML _processXml;
        INautilusDBConnection _ntlsCon;


        private INautilusServiceProvider _sp;

        private OracleConnection _connection;

        private OracleCommand cmd;

        private INautilusRecordSet rs;


        private List<ALIQUOT> aliquots;

        private double sessionId;

        private string _connectionString;

        #endregion

        #region Implementing IExtensionWindow
        public bool CloseQuery()
        {
            try
            {


                if (cmd != null)
                    cmd.Dispose();
                if (_connection != null)
                    _connection.Close();

                return true;
            }
            catch (Exception ex)
            {

                return true;
            }
        }

        public WindowRefreshType DataChange()
        {
            return LSExtensionWindowLib.
                WindowRefreshType.windowRefreshNone;
        }

        public WindowButtonsType GetButtons()
        {
            return LSExtensionWindowLib.WindowButtonsType.windowButtonsNone;
        }

        public void Internationalise()
        {
        }

        public void PreDisplay()
        {
            INautilusDBConnection dbConnection;
            if (_sp != null)
            {

                // Debugger.Launch();
                dbConnection = _sp.QueryServiceProvider("DBConnection") as NautilusDBConnection;
                rs = _sp.QueryServiceProvider("RecordSet") as NautilusRecordSet;



            }
            else
            {
                dbConnection = null;
            }



            _connection = GetConnection(dbConnection);
        }

        public void RestoreSettings(int hKey)
        {
        }

        public bool SaveData()
        {
            return false;
        }

        public void SaveSettings(int hKey)
        {
        }

        public void SetParameters(string parameters)
        {
        }

        public void SetServiceProvider(object serviceProvider)
        {
            _sp = serviceProvider as NautilusServiceProvider;

            if (_sp != null)
            {
                _processXml = _sp.QueryServiceProvider("ProcessXML") as NautilusProcessXML;
                _ntlsUser = _sp.QueryServiceProvider("User") as NautilusUser;
                _ntlsCon = Utils.GetNtlsCon(_sp);
            }
            else
            {
                _processXml = null;
            }
        }

        public void SetSite(object site)
        {
            _ntlsSite = (IExtensionWindowSite2)site;
            _ntlsSite.SetWindowInternalName("Fire Event");
            _ntlsSite.SetWindowRegistryName("Fire Event");
            _ntlsSite.SetWindowTitle("Fire Event");
        }

        public void Setup()
        {
        }

        public WindowRefreshType ViewRefresh()
        {
            return LSExtensionWindowLib.WindowRefreshType.windowRefreshNone;
        }

        public void refresh()
        {
        }

        #endregion


        #region Private methods

        public OracleConnection GetConnection(INautilusDBConnection ntlsCon)
        {

            OracleConnection connection = null;

            if (ntlsCon != null)
            {


                // Initialize variables
                String roleCommand;
                // Try/Catch block
                try
                {


                    var C = ntlsCon.GetServerIsProxy();
                    var C2 = ntlsCon.GetServerName();
                    var C4 = ntlsCon.GetServerType();

                    var C6 = ntlsCon.GetServerExtra();

                    var C8 = ntlsCon.GetPassword();
                    var C9 = ntlsCon.GetLimsUserPwd();
                    var C10 = ntlsCon.GetServerIsProxy();
                    var DD = _ntlsSite;




                    var u = _ntlsUser.GetOperatorName();
                    var u1 = _ntlsUser.GetWorkstationName();



                    _connectionString = ntlsCon.GetADOConnectionString();

                    var splited = _connectionString.Split(';');

                    var cs = "";

                    for (int i = 1; i < splited.Count(); i++)
                    {
                        cs += splited[i] + ';';
                    }
                    //<<<<<<< .mine
                    var username = ntlsCon.GetUsername();
                    if (string.IsNullOrEmpty(username))
                    {
                        var serverDetails = ntlsCon.GetServerDetails();
                        cs = "User Id=/;Data Source=" + serverDetails + ";";
                    }


                    //Create the connection
                    connection = new OracleConnection(cs);



                    // Open the connection
                    connection.Open();

                    // Get lims user password
                    string limsUserPassword = ntlsCon.GetLimsUserPwd();

                    // Set role lims user
                    if (limsUserPassword == "")
                    {
                        // LIMS_USER is not password protected
                        roleCommand = "set role lims_user";
                    }
                    else
                    {
                        // LIMS_USER is password protected.
                        roleCommand = "set role lims_user identified by " + limsUserPassword;
                    }

                    // set the Oracle user for this connecition
                    OracleCommand command = new OracleCommand(roleCommand, connection);

                    // Try/Catch block
                    try
                    {
                        // Execute the command
                        command.ExecuteNonQuery();
                    }
                    catch (Exception f)
                    {
                        // Throw the exception
                        throw new Exception("Inconsistent role Security : " + f.Message);
                    }

                    // Get the session id
                    sessionId = ntlsCon.GetSessionId();

                    // Connect to the same session
                    string sSql = string.Format("call lims.lims_env.connect_same_session({0})", sessionId);

                    // Build the command
                    command = new OracleCommand(sSql, connection);

                    // Execute the command
                    command.ExecuteNonQuery();

                }
                catch (Exception e)
                {
                    // Throw the exception
                    throw e;
                }

                // Return the connection
            }

            return connection;

        }
        #endregion

        private void aliquotName_KeyPress(object sender, KeyPressEventArgs e)
        {
            string sql = "";
            try
            {
                if (e.KeyChar == (char)13 && txtEditEntity.Text != "") //Enter
                {

                    //Checks if it's already in list view
                    if (!ListViewContains())
                    {
                        //Build query
                        if (!string.IsNullOrEmpty(txtEditEntity.Text))
                        {
                            sql = "select erd.U_SLIDE_NAME from lims_sys.u_extra_request_data_user erd inner join lims_sys.u_extra_request er " +
                                "on erd.U_EXTRA_REQUEST_ID = er.U_EXTRA_REQUEST_ID " +
                                "where er.name like '%Rescan%' and u_status='V' and U_REQ_TYPE='S' and erd.U_SLIDE_NAME = '" + txtEditEntity.Text +"'";
                        }
                        
                        cmd = new OracleCommand(sql, _connection);
                        OracleDataReader reader = cmd.ExecuteReader();

                        //Checks if it exists
                        if (!reader.HasRows)
                        {
                            MessageBox.Show(txtEditEntity.Text +
                                "  does not exist or does not meet the conditions!", "Nautilus",
                                MessageBoxButtons.OK, MessageBoxIcon.Hand);
                            txtEditEntity.Focus();
                        }
                        else
                        {
                            ListViewItem li = null;
                            while (reader.Read())
                            {
                                li = new ListViewItem(reader[0].ToString());
                                string name = reader[0].ToString();
                            }
                            listViewEntities.Items.Add(li);
                            txtEditEntity.Text = string.Empty;
                            reader.Close();
                            listViewEntities.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
                        }
                    }
                    else
                    {
                        MessageBox.Show(txtEditEntity.Text + "  already exists");
                        txtEditEntity.Focus();
                    }
                }
            }
            catch (Exception e1)
            {
                MessageBox.Show("Error: " + e1.Message + e1.StackTrace);
                //Logger.WriteLogFile(e1);
            }
        }

        private bool ListViewContains()
        {
            foreach (ListViewItem item in listViewEntities.Items)
            {
                if (item.SubItems[0].Text == txtEditEntity.Text)
                    return true;
            }
            return false;
        }

        private void Ok_button_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (ListViewItem item in listViewEntities.Items)
                {
                    RunEvent(item.SubItems[0].Text);
                    Add2Log(item.SubItems[0].Text);
                }
                //Empties the list
                listViewEntities.Clear();
                MessageBox.Show("Completed",this.Name,MessageBoxButtons.OK);
            }
            catch (Exception e1)
            {
                MessageBox.Show("Error" + e1.Message);

                //Logger.WriteLogFile(e1);
            }
        }

        //event needs to be implemented later
        private void RunEvent(string name)
        {
            string sql = "select u_slide_name from lims_sys.U_EXTRA_REQUEST_DATA_USER where U_REQ_TYPE='S' and   u_status='V' and U_SLIDE_NAME = '" + name + "'";
            cmd = new OracleCommand(sql, _connection);
            OracleDataReader reader = cmd.ExecuteReader();

            //Checks if it exists
            if (!reader.HasRows)
            {
                MessageBox.Show(name +
                    "  does not exist or does not meet the conditions!", "Nautilus",
                    MessageBoxButtons.OK, MessageBoxIcon.Hand);
                txtEditEntity.Focus();
            }
            else
            {
                while (reader.Read())
                {
                    string ex = reader[0].ToString();
                    sql = "update lims_sys.U_EXTRA_REQUEST_DATA_USER set u_status = 'L', U_LAB_ON = to_char(sysdate) where U_REQ_TYPE='S' and   u_status='V' and U_SLIDE_NAME = '" + ex + "'";
                    cmd = new OracleCommand(sql, _connection);
                    cmd.ExecuteReader();
                }
            }
            reader.Close();

        }

        //log needs to be implemented later
        private void Add2Log(string id)
        {
        }

        private void close_button_Click(object sender, EventArgs e)
        {
            try
            {
                if (listViewEntities.Items.Count > 0)
                {
                    DialogResult dialogResult = MessageBox.Show("האם אתה בטוח שברצונך לצאת ממסך זה ללא אישור? ", "יציאה",
                        MessageBoxButtons.YesNoCancel);
                    if (dialogResult == DialogResult.Yes)
                    {
                        listViewEntities = null;
                        if (_ntlsSite != null)
                            _ntlsSite.CloseWindow();
                    }
                }
                else
                {
                    listViewEntities = null;
                    if (_ntlsSite != null)
                        _ntlsSite.CloseWindow();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

    }
    
}
