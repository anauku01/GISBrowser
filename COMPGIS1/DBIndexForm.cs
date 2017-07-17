using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading;

namespace COMPGIS1
{

    public partial class BPSettingsForm : Form
    {

        private ArcEngineClass _GISdata;
        private ArcEngineBP _BPGISdata;
        private const int su_layerLines = 1;
        private const int su_layerLineSections = 2;
        private const int su_layerBoringLocations = 3;
        private const int su_layerCorrosionProbes = 4;
        private bool processingsynch = false;
        private List<string> _layerlist = new List<string>();

        //---------------------------------------------------
        // Constructor
        //---------------------------------------------------
        public BPSettingsForm(ArcEngineBP BPGISdata)
        {
            InitializeComponent();
            _BPGISdata = BPGISdata;
            _GISdata = (ArcEngineClass) BPGISdata;
        }

        //---------------------------------------------------
        // Form Load Event
        //---------------------------------------------------
        private void DBSynchForm_Load(object sender, EventArgs e)
        {
            SetControls();
        }

        //---------------------------------------------------
        // Sets controls based upon form control settings
        //---------------------------------------------------
        private void SetControls()
        {
            cbLines.Enabled = (_GISdata != null);
            cbLineSections.Enabled = cbLines.Enabled;
            cbBoringLocations.Enabled = cbLines.Enabled;
            cbCorrosionProbes.Enabled = cbLines.Enabled;
            btnReIndex.Enabled = (cbLines.Enabled && (cbLines.Checked || cbLineSections.Checked || cbCorrosionProbes.Checked || cbBoringLocations.Checked) || (cbRectifiers.Checked) || cbTestStations.Checked);
        }

        //---------------------------------------------------
        // Show the progress callback
        //---------------------------------------------------
        private int UpdateProgress(int totalrecs, int currentrec)
        {
            return 0;
        }


        //---------------------------------------------------
        // Run the Synch
        //---------------------------------------------------
        private void RunSynch()
        {

            btnClose.Enabled = false;
            tabSynchronize.Enabled = false;
            lblStatus.Text = string.Format("Please wait: {0}% Completed", 0);
            lblStatus.Update();
            lblStatus.Update();
            processingsynch = true;
            string currenttablename = "";

            BackgroundWorker bw = new BackgroundWorker();
            // this allows our worker to report progress during work
            bw.WorkerReportsProgress = true;
            lblStatus.Visible = true;
            int TotalToDo = 0;
            if (cbLines.Checked)
            {
                TotalToDo++;
                currenttablename = cbLines.Text;
            }               
            if (cbLineSections.Checked)
            {
                TotalToDo++;
                if (currenttablename.Length == 0)
                    currenttablename = cbLineSections.Text;
            }
            if (cbBoringLocations.Checked)
            {
                TotalToDo++;
                if (currenttablename.Length == 0)
                    currenttablename = cbBoringLocations.Text;
            }
            if (cbCorrosionProbes.Checked)
            {
                TotalToDo++;
                if (currenttablename.Length == 0)
                    currenttablename = cbCorrosionProbes.Text;
            }
            if (cbRectifiers.Checked)
            {
                TotalToDo++;
                if (currenttablename.Length == 0)
                    currenttablename = cbRectifiers.Text;
            }
            if (TotalToDo == 0) return;

            lblStatus2.Text = string.Format("Processing: {0}",currenttablename);

            // what to do in the background thread
            bw.DoWork += new DoWorkEventHandler(
            delegate(object o, DoWorkEventArgs args)
            {
                BackgroundWorker b = o as BackgroundWorker;
                // process each of the layers
                int j = 0;
                for (int i = 1; i <= 5; i++)
                {
                    double pcent = ((double) j/ (double) TotalToDo)*100;
                    int pcint = (int) Math.Truncate(pcent);
                    switch (i)
                    {
                        case 1:
                            {
//                                lblStatus.Text = "Synchronizing Lines";
                                if (cbLines.Checked)
                                {
                                    _GISdata._cwdata.lineClearRecordCrossReferences();
                                    _BPGISdata.linesUpdateCrossReference();
                                    j++;
                                }
                                break;
                            }
                        case 2 :
                            {
//                                lblStatus.Text = "Synchronizing Line Sections";
                                if (cbLineSections.Checked)
                                {
                                    _GISdata._cwdata.linesectionClearRecordCrossReferences();
                                    _BPGISdata.linesectionUpdateCrossReference();
                                    currenttablename = cbLineSections.Text;
                                    j++;
                                }
                                break;
                            }

                        case 3 :
                            {
//                                lblStatus.Text = "Synchronizing Boring Locations";
                                if (cbBoringLocations.Checked)
                                {
                                    _GISdata._cwdata.blClearRecordCrossReferences();
                                    _BPGISdata.boringlocationsUpdateCrossReference();
                                    currenttablename = cbBoringLocations.Text;
                                    j++;
                                }
                                break;
                            }

                        case 4 :
                            {
//                                lblStatus.Text = "Synchronizing Corrosion Probes";
                                if (cbCorrosionProbes.Checked)
                                {
                                    _GISdata._cwdata.cpClearRecordCrossReferences();
                                    _BPGISdata.corrosionprobesUpdateCrossReference();
                                    currenttablename = cbCorrosionProbes.Text;
                                    j++;
                                }
                                break;
                            }

                        case 5:
                            {
                                if (cbRectifiers.Checked)
                                {
                                    _GISdata._cwdata.rectClearRecordCrossReferences();
                                    _BPGISdata.rectifierUpdateCrossReference();
                                    currenttablename = cbRectifiers.Text;
                                    j++;
                                }
                                break;
                            }

                        case 6:
                            {
                                if (cbTestStations.Checked)
                                {
                                    _GISdata._cwdata.teststationClearRecordCrossReferences();
                                    _BPGISdata.layerUpdateCrossReference(LayerSystemUse.suBPTestStations);
                                    currenttablename = cbTestStations.Text;
                                    j++;
                                }
                                break;
                            }
                    
                    
                    }
                    // report the progress in percent
                    b.ReportProgress(pcint);
                    Thread.Sleep(500);
                }
                b.ReportProgress(100);

            });

            // what to do when progress changed (update the progress bar for example)
            bw.ProgressChanged += new ProgressChangedEventHandler(
            delegate(object o, ProgressChangedEventArgs args)
            {
                lblStatus.Text = string.Format("Please wait: {0}% Completed", args.ProgressPercentage);
                lblStatus.Update();
                lblStatus2.Text = string.Format("Processing: {0}", currenttablename);
                lblStatus2.Update();
            });

            // what to do when worker completes its task (notify the user)
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(
            delegate(object o, RunWorkerCompletedEventArgs args)
            {
                lblStatus.Text = "";
                lblStatus2.Text = "";
                lblStatus.Update();
                tabSynchronize.Enabled = true;
                // cleanup
                processingsynch = false;
                btnClose.Enabled = true;
            });


            // Run the synch
            bw.RunWorkerAsync();
        }


        //---------------------------------------------------
        // Do the Indexing
        //---------------------------------------------------
        private void btnReIndex_Click(object sender, EventArgs e)
        {
            RunSynch();
        }

        //---------------------------------------------------
        // One of the Checkboxes has changed
        //---------------------------------------------------
        private void cbLines_Click(object sender, EventArgs e)
        {
            SetControls();
        }

        //---------------------------------------------------
        // Close the Form
        //---------------------------------------------------
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //---------------------------------------------------
        // Form Close Event
        //---------------------------------------------------
        private void DBSynchForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Do not allow close if the form is processing synch
            e.Cancel = processingsynch;
        }

        //---------------------------------------------------
        // Double click on the Key Map grid
        //---------------------------------------------------
        private void gridKeyFields_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
        }




    }
}


