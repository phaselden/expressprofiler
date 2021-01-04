using System;
using System.Windows.Forms;

namespace ExpressProfiler
{
    public partial class FindForm : Form
    {
        private MainForm m_mainForm;

        public FindForm(MainForm f)
        {
            InitializeComponent();

            m_mainForm = f;

            // Set the control values to the last find performed.
            edPattern.Text = m_mainForm._lastPattern;
            chkCase.Checked = m_mainForm._matchCase;
            chkWholeWord.Checked = m_mainForm._wholeWord;
        }

        private void btnFindNext_Click(object sender, EventArgs e)
        {
            DoFind(true);
        }

        private void btnFindPrevious_Click(object sender, EventArgs e)
        {
            DoFind(false);
        }

        private void DoFind(bool forwards)
        {
            m_mainForm._lastPattern = edPattern.Text;
            m_mainForm._matchCase = chkCase.Checked;
            m_mainForm._wholeWord = chkWholeWord.Checked;
            m_mainForm.PerformFind(forwards, chkWrapAround.Checked);
        }

        private void edPattern_TextChanged(object sender, EventArgs e)
        {
            m_mainForm._lastPos = -1;
        }
    }
}
