using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportExcelToDataGrid
{
    public partial class FrmImportExcelToDatagrid : Form
    {
        //-------------------------------------------------------------------------------------------------------------
        //                                                  //INSTANCE VARIABLES.
        private String path = "C:\\Users\\Cesar Garcia\\source\\reposFormCSharp\\YOUTUBE-CSharpImportExcelToDataGrid\\ImportExcelToDataGrid\\excelTest.xlsx";

        //-------------------------------------------------------------------------------------------------------------
        //                                                  //CONSTRUCTORS.
        public FrmImportExcelToDatagrid()
        {
            InitializeComponent();
        }

        //--------------------------------------------------------------------------------------------------------------
        //                                                  //ACCESS METHODS.
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //--------------------------------------------------------------------------------------------------------------
        private void FrmImportExcelToDatagrid_Load(object sender, EventArgs e)
        {
            SLDocument sl = new SLDocument(path);

            List<PersonasViewModel> darrperviemodelPersonasViewModel = new List<PersonasViewModel>();
            int intRow = 2;
            /*REPEAT-WHILE*/
            while (
                //
                !String.IsNullOrEmpty(sl.GetCellValueAsString(intRow, 1))
                )
            {
                PersonasViewModel prvm = new PersonasViewModel();
                prvm.strCode = sl.GetCellValueAsString(intRow, 1);
                prvm.strName = sl.GetCellValueAsString(intRow, 2);
                prvm.intEdad = sl.GetCellValueAsInt32(intRow, 2);

                darrperviemodelPersonasViewModel.Add(prvm);
                intRow = intRow + 1;
            }

            dataGridView1.DataSource = darrperviemodelPersonasViewModel;
        }
    }
}
