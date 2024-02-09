namespace DocxEasyFormat
{
    partial class EditorsFriend
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EditorsFriend));
            label2 = new Label();
            listView1 = new ListView();
            Filename = new ColumnHeader();
            Status = new ColumnHeader();
            SuspendLayout();
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(8, 13);
            label2.MaximumSize = new Size(525, 0);
            label2.Name = "label2";
            label2.Size = new Size(500, 30);
            label2.TabIndex = 3;
            label2.Text = "A small utility to modify DOCX files. This will set the line space to single, remove any padding before and after paragraphs, set the font to Calibri, and set the font-size to 11. ";
            // 
            // listView1
            // 
            listView1.AllowDrop = true;
            listView1.Columns.AddRange(new ColumnHeader[] { Filename, Status });
            listView1.Location = new Point(8, 46);
            listView1.Name = "listView1";
            listView1.Size = new Size(525, 223);
            listView1.TabIndex = 4;
            listView1.UseCompatibleStateImageBehavior = false;
            listView1.View = View.Details;
            listView1.DragDrop += panel1_DragDrop;
            listView1.DragEnter += panel1_DragEnter;
            // 
            // Filename
            // 
            Filename.Text = "File Name";
            Filename.Width = 425;
            // 
            // Status
            // 
            Status.Text = "Status";
            Status.Width = 428;
            // 
            // EditorsFriend
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoScroll = true;
            ClientSize = new Size(543, 276);
            Controls.Add(listView1);
            Controls.Add(label2);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(2);
            Name = "EditorsFriend";
            SizeGripStyle = SizeGripStyle.Hide;
            Text = "Editor's Friend";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Label label2;
        private ListView listView1;
        private ColumnHeader Filename;
        private ColumnHeader Status;
    }
}