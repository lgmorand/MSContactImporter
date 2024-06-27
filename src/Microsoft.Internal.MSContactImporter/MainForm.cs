using Microsoft.Internal.MSContactImporter.Properties;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;

namespace Microsoft.Internal.MSContactImporter
{
    public partial class MainForm : Form
    {
        private List<RootMSFTee> rootMSFTees;

        public MainForm()
        {
            InitializeComponent();

            checkBoxImportPhotos.Checked = Settings.Default.ImportPhotos;
        }

        #region Events

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Original idea: Pascal B. (a.k.a Président Maréchal of FrFreeD)\r\nUpgraded by: Louis-Guillaume M. (a.k.a LG)\r\n" +
                "Upgraded by: Aurélien N. (a.k.a aurnor)\r\n" +
                "Sponsored by La e-cig c'est mal\r\nUsed by Votez FrFreed sur Yammer en 2076\r\nValidated by Qui s'occupe des stages ?");
        }

        private async void btnGo_Click(object sender, EventArgs e)
        {
            this.SaveRootMSFTees();
            string operation = string.Empty;

            Stopwatch sw = new Stopwatch();
            sw.Start();
            if (rdioImport.Checked)
            {
                operation = "import";
                await ImportContacts();
            }
            else if (rdoUpdate.Checked)
            {
                operation = "update contacts";
                await UpdateContacts();
            }
            else if (rdioDelete.Checked)
            {
                operation = "cleaning orphans contacts";
                DeleteOrphansContacts();
            }
            sw.Stop();

            TimeSpan ts = sw.Elapsed;
            Logger.LogMessageToConsole($"Operation took: {ts.Hours}h {ts.Minutes}min {ts.Seconds}sec {ts.Milliseconds}ms");

            if (checkBoxImportPhotos.Checked)
                ClearPhotoCache();

            MessageBox.Show(operation + " is finished");
            OfferToDeleteLogs();
        }

        private void OfferToDeleteLogs()
        {
            if(MessageBox.Show("As the technical logs may contain personal information like the names and email of imported/edited contacts, we recommend to delete them. Do you want delete those files ?","GDPR", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                try
                {
                    File.Delete("log.txt");
                }
                catch
                {
                    MessageBox.Show("We were unable to delete the logs files located in the same folder than the app. Please delete them manually");
                }
            }
        }

        private void ClearPhotoCache()
        {
            DirectoryInfo di = new DirectoryInfo(Application.UserAppDataPath + "\\photos");
            if (di.Exists)
            {
                di.Delete(true);
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = tabControl.SelectedIndex + 1;
            btnNext.Enabled = tabControl.SelectedIndex != tabControl.TabCount - 1;
            btnPrevious.Visible = tabControl.SelectedIndex != 0;
            btnPrevious.Enabled = tabControl.SelectedIndex != 0;
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            tabControl.SelectedIndex = tabControl.SelectedIndex - 1;
            btnNext.Enabled = tabControl.SelectedIndex != tabControl.TabCount - 1;
            btnPrevious.Visible = tabControl.SelectedIndex != 0;
            btnPrevious.Enabled = tabControl.SelectedIndex != 0;

            btnPrevious.Visible = tabControl.SelectedIndex != 0;

            btnTest.Visible = tabControl.SelectedIndex == 0;
            btnTest.Enabled = tabControl.SelectedIndex == 0;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            btnTest.Enabled = false;

            using (OutlookUtils outlookUtils = new OutlookUtils())
            {
                if (outlookUtils.TestConnection())
                {
                    if (new ADUtils(txtEmail.Text, txtPassword.Text).TestConnection(this.rootMSFTees.First().Logon))
                    {
                        btnTest.Visible = false;
                        btnNext_Click(btnNext, e);
                    }
                }
            }
            btnTest.Enabled = true;
        }


        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Logger.SetControl(txtConsole);

            txtEmail.Text = GetLoggedUser();

            this.Text = string.Format("{0} (v {1})", this.Text, Assembly.GetExecutingAssembly().GetName().Version);

            GetHierarchy();

            PrepareGridView();
        }

        #endregion Events

        #region Helpers

        internal XElement ExportRootMSFTeesToXml()
        {
            return new XElement("RootMSFTees", this.rootMSFTees.Select(m => ExportRootMSFTeeToXml(m)));
        }

        internal XElement ExportRootMSFTeeToXml(RootMSFTee rootMSFTee)
        {
            return rootMSFTee.ToXml();
        }

        private void SaveRootMSFTees()
        {
            Settings.Default.RootMSFTees = this.ExportRootMSFTeesToXml().ToString();
            Settings.Default.Save();
        }

        internal DialogResult HandleAbortRetryCancel(Exception exception, IState state)
        {
            return MessageBox.Show(this, string.Format("An error has occured:\r\n\nModule: {1}\r\nException: {0}\r\n\r\nWhat do you want to do?", exception.Message, ((State)state).Function), "Error", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1);
        }

        internal RootMSFTee XmlToRootMsfte(XElement xElement)
        {
            return RootMSFTee.FromXml(xElement);
        }

        private double CalculatePercentage(int value, int max)
        {
            return Convert.ToDouble(value) / Convert.ToDouble(max);
        }

        #endregion Helpers

        private void DeleteOrphansContacts()
        {
            Logger.LogMessageToConsole("Starting 'Deleting Contacts' operation...");

            int ignored = 0;

            int deleted = 0;

            try
            {
                bool cancelDeleteContacts = false;
                ADUtils adUtils = new ADUtils(txtEmail.Text, txtPassword.Text);
                Logger.LogMessageToConsole("Loading existing contacts...");

                using (OutlookUtils outlookUtils = new OutlookUtils())
                {

                    State state = new State
                    {
                        Function = "LoadContacts"
                    };
                    AbortRetryIgnorePattern.CallMethod(delegate
                    {
                        outlookUtils.LoadContacts();
                    }, state, delegate
                    {
                        cancelDeleteContacts = true;
                    }, new Func<Exception, IState, DialogResult>(HandleAbortRetryCancel));
                    if (cancelDeleteContacts)
                    {
                        Logger.LogMessageToConsole("Operation cancelled");
                        this.progressBar.Value = this.progressBar.Maximum;

                        Logger.LogMessageToConsole(string.Format("Number of contacts ignored: {0}", ignored));
                        Logger.LogMessageToConsole(string.Format("Number of contacts deleted: {0}", deleted));
                        Logger.LogMessageToConsole("'Deleting Contacts' operation completed");
                        return;
                    }
                    Logger.LogMessageToConsole("Loading contacts completed!");
                    this.progressBar.Minimum = 0;
                    this.progressBar.Maximum = outlookUtils.Contacts.Count;
                    this.progressBar.Value = 0;

                    List<Outlook.ContactItem> contactsToRemove = new List<Outlook.ContactItem>();
                    outlookUtils.Contacts.ForEach(delegate (Outlook.ContactItem contact)
                    {
                        if (cancelDeleteContacts)
                        {
                            Logger.LogMessageToConsole("Operation cancelled");
                            progressBar.Value = progressBar.Maximum;
                            return;
                        }

                        state = new State
                        {
                            Function = "GetDistinguishedName"
                        };

                        string distinguishedName = AbortRetryIgnorePattern.CallFunction<string>(delegate
                        {
                            Outlook.PropertyAccessor pa = contact.PropertyAccessor;
                            string alias = pa.GetProperty(Settings.Default.ExtendedPropertySchema + Settings.Default.MsStaffId) as string;
                            Marshal.ReleaseComObject(pa);
                            return adUtils.GetDistinguishedName(alias);
                        }, state, delegate
                        {
                            cancelDeleteContacts = true;
                        }, new Func<Exception, IState, DialogResult>(HandleAbortRetryCancel));

                        if (cancelDeleteContacts)
                        {
                            Logger.LogMessageToConsole("Operation cancelled");

                            progressBar.Value = progressBar.Maximum;

                            return;
                        }

                        if (string.IsNullOrEmpty(distinguishedName))
                        {
                            Logger.LogMessageToConsole(string.Format("{0:#00%} - Deleting contact: {1}", CalculatePercentage(progressBar.Value, progressBar.Maximum), contact.FullName));

                            state = new State
                            {
                                Function = "Delete"
                            };

                            AbortRetryIgnorePattern.CallMethod(delegate
                            {
                                contactsToRemove.Add(contact);

                            }, state, delegate
                            {
                                cancelDeleteContacts = true;
                            }, new Func<Exception, IState, DialogResult>(HandleAbortRetryCancel));

                            if (cancelDeleteContacts)
                            {
                                Logger.LogMessageToConsole("Operation cancelled");

                                progressBar.Value = progressBar.Maximum;

                                return;
                            }
                            deleted += 1;

                            Logger.LogMessageToConsole(string.Format("Contact deleted: {0}", contact.FullName));
                        }
                        else
                        {
                            Logger.LogMessageToConsole(string.Format("{0:#00%} - Ignoring contact: {1}", CalculatePercentage(progressBar.Value, progressBar.Maximum), contact.FullName));
                            ignored += 1;

                            Logger.LogMessageToConsole(string.Format("Contact ignored: {0}", contact.FullName));
                        }

                        progressBar.Value++;
                    });
                    if (cancelDeleteContacts)
                    {
                        Logger.LogMessageToConsole("Operation cancelled!");
                    }

                    foreach (var contact in contactsToRemove)
                    {
                        contact.Delete();
                        Marshal.ReleaseComObject(contact);
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.ToString(), "Error - buttonDeleteContacts_Click", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }

            this.progressBar.Value = this.progressBar.Minimum;

            Logger.LogMessageToConsole(string.Format("Number of contacts ignored: {0}", ignored));
            Logger.LogMessageToConsole(string.Format("Number of contacts deleted: {0}", deleted));
            Logger.LogMessageToConsole("'Deleting Contacts' operation completed");
        }

        private void GetHierarchy()
        {
            try
            {
                IEnumerable<XElement> users = XElement.Parse(Settings.Default.RootMSFTees).Elements("RootMSFTee");
                this.rootMSFTees = users.Select(XmlToRootMsfte).ToList<RootMSFTee>();
            }
            catch
            {
                this.rootMSFTees = new List<RootMSFTee>
                {
                    new RootMSFTee
                    {
                        Logon = "capurass",
                        RecurseLevel = 7
                    }
                };
                Settings.Default.RootMSFTees = ExportRootMSFTeesToXml().ToString();
                Settings.Default.Save();
            }
        }

        /// <summary>
        /// Retrieves the current user
        /// </summary>
        private string GetLoggedUser()
        {
            if (string.IsNullOrEmpty(Settings.Default.Mailbox))
            {
                Settings.Default.Mailbox = string.Format("{0}@microsoft.com", Environment.UserName);
                Settings.Default.Save();
            }
            return Settings.Default.Mailbox;
        }

        private async Task ImportContacts()
        {
            txtConsole.Text = string.Empty;
            Logger.LogMessageToConsole("Starting 'Importing Contacts' operation...");

            int insertedContacts = 0;
            int updatedContacts = 0;
            State state = null;
            bool cancelImportContacts = false;


            try
            {
                GraphUtils graphUtils = null;
                if (checkBoxImportPhotos.Checked)
                {
                    graphUtils = new GraphUtils();
                    await graphUtils.SigninAsync();
                }

                using (OutlookUtils outlookUtils = new OutlookUtils())
                {
                    Logger.LogMessageToConsole("Starting loading contacts...");
                    state = new State
                    {
                        Function = "LoadContacts"
                    };

                    AbortRetryIgnorePattern.CallMethod(delegate
                                {
                                    outlookUtils.LoadContacts();
                                }, state, delegate
                                {
                                    cancelImportContacts = true;
                                }, new Func<Exception, IState, DialogResult>(HandleAbortRetryCancel));

                    if (cancelImportContacts)
                    {
                        Logger.LogMessageToConsole("Operation cancelled!");
                        return;
                    }

                    Logger.LogMessageToConsole("Loading your contacts completed");

                    ADUtils adUtils = new ADUtils(txtEmail.Text, txtPassword.Text);

                    this.rootMSFTees.ForEach(delegate (RootMSFTee rootMSFTee)
                    {
                        Logger.LogMessageToConsole(string.Format("Loading MSFTee tree for {0}...", rootMSFTee.Logon));

                        state = new State
                        {
                            Function = "LoadActiveDirectory"
                        };

                        Action loadAdForRootUser = delegate
                        {
                            Func<bool> cancelOperation = delegate
                            {
                                return false;
                            };

                            Action<string> log = delegate (string label)
                            {
                                Logger.LogMessageToConsole(string.Format("Root: {0} - Total: {2:### ##0}\r\nLoading information for: {1}...", rootMSFTee.Logon, label, adUtils.MSFTees.Count));
                            };

                            adUtils.LoadActiveDirectory(rootMSFTee.Logon, rootMSFTee.RecurseLevel != 0, rootMSFTee.RecurseLevel, cancelOperation, log);
                        };

                        AbortRetryIgnorePattern.CallMethod(loadAdForRootUser,
                            state,
                            delegate
                            {
                                cancelImportContacts = true;
                            },
                            new Func<Exception, IState, DialogResult>(HandleAbortRetryCancel));
                    });

                    if (cancelImportContacts)
                    {
                        this.progressBar.Value = this.progressBar.Maximum;
                        Logger.LogMessageToConsole("Operation cancelled!");
                        return;
                    }
                    Logger.LogMessageToConsole(string.Format("{0} msftees found to import/update", adUtils.MSFTees.Count));
                    this.progressBar.Minimum = 0;
                    this.progressBar.Maximum = adUtils.MSFTees.Count;
                    this.progressBar.Value = 0;

                    adUtils.MSFTees.Values.ToList<MSFTee>().ForEach(delegate (MSFTee msftee)
                    {
                        if (cancelImportContacts)
                        {
                            Logger.LogMessageToConsole("Operation cancelled");
                            this.progressBar.Value = progressBar.Maximum;
                            return;
                        }

                        state = new State
                        {
                            Function = "InsertOrUpdateContact"
                        };

                        Func<bool> insert = () => outlookUtils.InsertOrUpdateContact(msftee, graphUtils);

                        bool flag = AbortRetryIgnorePattern.CallFunction<bool>(insert, state,
                        delegate
                        {
                            cancelImportContacts = true;
                        }, new Func<Exception, IState, DialogResult>(HandleAbortRetryCancel));

                        if (cancelImportContacts)
                        {
                            Logger.LogMessageToConsole("Operation cancelled");
                            progressBar.Value = progressBar.Maximum;
                            return;
                        }

                        progressBar.Value++;

                        if (flag)
                        {
                            insertedContacts += 1;
                            Logger.LogMessageToConsole(string.Format("Contact inserted: {0}", msftee.FullName));
                            Logger.LogMessageToConsole(string.Format("{0:#00%} - MSFTee updated: {1}", CalculatePercentage(this.progressBar.Value, progressBar.Maximum), msftee.FullName));
                        }
                        else
                        {
                            updatedContacts += 1;
                            Logger.LogMessageToConsole(string.Format("Contact updated: {0}", msftee.FullName));
                            Logger.LogMessageToConsole(string.Format("{0:#00%} - MSFTee updated: {1}", CalculatePercentage(this.progressBar.Value, progressBar.Maximum), msftee.FullName));
                        }
                    });

                    if (cancelImportContacts)
                    {
                        Logger.LogMessageToConsole("Operation cancelled!");
                    }

                    this.progressBar.Value = this.progressBar.Minimum;
                    Logger.LogMessageToConsole(string.Format("Number of contacts inserted: {0}", insertedContacts));
                    Logger.LogMessageToConsole(string.Format("Number of contacts updated: {0}", updatedContacts));
                    Logger.LogMessageToConsole("'Importing Contacts' operation completed");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Error - buttonImportContacts_Click", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void linkUrl_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://kb.intermedia.net/article/2150");
        }

        private void LoadDataGridMSFTees()
        {
            this.gdvSettings.AutoGenerateColumns = false;
            this.gdvSettings.DataSource = null;
            this.gdvSettings.DataSource = this.rootMSFTees;
        }

        private void PrepareGridView()
        {
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(-1, "All"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(0, "None"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(1, "One level"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(2, "Two levels"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(3, "Three levels"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(4, "Four levels"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(5, "Five levels"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(6, "Six levels"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(7, "Seven levels"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(8, "Eight levels"));
            this.RecurseLevel.Items.Add(new KeyValuePair<int, string>(9, "Nine levels"));
            this.RecurseLevel.ValueMember = "Key";
            this.RecurseLevel.DisplayMember = "Value";

            LoadDataGridMSFTees();
        }

        private async Task UpdateContacts()
        {
            Logger.LogMessageToConsole("Starting 'Updating Contacts' operation...");

            int updated = 0;

            try
            {
                GraphUtils graphUtils = null;
                if (checkBoxImportPhotos.Checked)
                {
                    graphUtils = new GraphUtils();
                    await graphUtils.SigninAsync();
                }

                bool cancelUpdateContacts = false;

                ADUtils adUtils = new ADUtils(txtEmail.Text, txtPassword.Text);

                using (OutlookUtils outlookUtils = new OutlookUtils())
                {
                    Logger.LogMessageToConsole("Starting loading contacts...");

                    State state = new State
                    {
                        Function = "LoadContacts"
                    };
                    AbortRetryIgnorePattern.CallMethod(delegate
                    {
                        outlookUtils.LoadContacts();
                    }, state, delegate
                    {
                        cancelUpdateContacts = true;
                    }, new Func<Exception, IState, DialogResult>(HandleAbortRetryCancel));

                    if (cancelUpdateContacts)
                    {
                        this.progressBar.Value = this.progressBar.Maximum;
                        Logger.LogMessageToConsole("'Updating Contacts' operation cancelled");
                        return;
                    }

                    Logger.LogMessageToConsole("Loading your contacts completed");
                    this.progressBar.Minimum = 0;
                    this.progressBar.Maximum = outlookUtils.Contacts.Count;
                    this.progressBar.Value = 0;



                    outlookUtils.Contacts.ForEach(delegate (Outlook.ContactItem contact)
                    {
                        if (cancelUpdateContacts)
                        {
                            Logger.LogMessageToConsole("Operation cancelled!");
                            progressBar.Value = progressBar.Maximum;
                            return;
                        }

                        Logger.LogMessageToConsole(string.Format("{0:#00%} - Updating contact: {1}", CalculatePercentage(progressBar.Value, progressBar.Maximum), contact.FullName));

                        state = new State
                        {
                            Function = "UpdateContact"
                        };

                        AbortRetryIgnorePattern.CallMethod(delegate
                        {
                            outlookUtils.UpdateContact(contact, adUtils, graphUtils);
                        }, state,
                        delegate
                        {
                            cancelUpdateContacts = true;
                        },
                        new Func<Exception, IState, DialogResult>(HandleAbortRetryCancel));

                        if (cancelUpdateContacts)
                        {
                            Logger.LogMessageToConsole("Operation cancelled!");
                            progressBar.Value = progressBar.Maximum;
                            return;
                        }

                        progressBar.Value++;

                        Logger.LogMessageToConsole(string.Format("Contact updated: {0}", contact.FullName));
                        Marshal.ReleaseComObject(contact);
                        updated += 1;
                    });

                    if (cancelUpdateContacts)
                    {
                        Logger.LogMessageToConsole("Operation cancelled!");
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(this, exception.ToString(), "Error - buttonUpdateContacts_Click", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }

            this.progressBar.Value = this.progressBar.Minimum;

            Logger.LogMessageToConsole(string.Format("Number of contacts updated: {0}", updated));
            Logger.LogMessageToConsole("'Updating Contacts' operation completed");
        }

        private void btnAddAlias_Click(object sender, EventArgs e)
        {
            int index = 1;
            while ((from r in this.rootMSFTees
                    where r.Logon == string.Format("[NewLogon{0}]", index)
                    select r).FirstOrDefault<RootMSFTee>() != null)
            {
                int index2 = index;
                index = index2 + 1;
            }
            this.rootMSFTees.Add(new RootMSFTee
            {
                Logon = string.Format("[NewLogon{0}]", index),
                RecurseLevel = -1
            });
            this.SaveRootMSFTees();
            this.LoadDataGridMSFTees();
        }

        private void btnDeleteAlias_Click(object sender, EventArgs e)
        {
            IEnumerable<DataGridViewRow> rows = this.gdvSettings.SelectedRows.OfType<DataGridViewRow>();

            rows.Select(r => r.Cells[0].Value.ToString()).ToList<string>().ForEach(delegate (string l)
            {
                this.rootMSFTees.Remove((from r in this.rootMSFTees
                                         where r.Logon == l
                                         select r).First<RootMSFTee>());
            });
            this.SaveRootMSFTees();
            this.LoadDataGridMSFTees();
        }

        private void checkBoxImportPhotos_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.ImportPhotos = checkBoxImportPhotos.Checked;
            Settings.Default.Save();
        }
    }
}