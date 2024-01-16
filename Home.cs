using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using System.Xml.Linq;


namespace ProvaN3_SolicitaSuporte
{
    public partial class Home : Form
    {
        public Home()
        {
            InitializeComponent();
        }

        

        private void btn_SendImage_Click(object sender, EventArgs e)
        {
            string todayDate = DateTime.Now.ToString("ddMMyyyyHHmmss");
            string fileName = $"Screenshot_{todayDate}.jpg";
            string filePath = Path.Combine(@"C:\teste\", fileName);
            string directory = Path.GetDirectoryName(filePath);

            Email newEmail = new Email();

            this.WindowState = FormWindowState.Minimized;
            Screenshot.TakeScreenshot(filePath, directory);
            this.WindowState = FormWindowState.Normal;

            newEmail.SendImageToSuport(txtBox_To.Text, txtBox_Subject.Text, txtBox_BodyEmail.Text, filePath);

        }

    }
    
}
