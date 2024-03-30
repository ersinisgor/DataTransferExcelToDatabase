namespace DataTransferExcelToDatabase
{
	partial class Form1
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
			DatabaseToExcel = new Button();
			richTextBox1 = new RichTextBox();
			SuspendLayout();
			// 
			// DatabaseToExcel
			// 
			DatabaseToExcel.Location = new Point(563, 66);
			DatabaseToExcel.Name = "DatabaseToExcel";
			DatabaseToExcel.Size = new Size(104, 23);
			DatabaseToExcel.TabIndex = 0;
			DatabaseToExcel.Text = "DatabaseToExcel";
			DatabaseToExcel.UseVisualStyleBackColor = true;
			DatabaseToExcel.Click += DatabaseToExcel_Click;
			// 
			// richTextBox1
			// 
			richTextBox1.Location = new Point(94, 66);
			richTextBox1.Name = "richTextBox1";
			richTextBox1.Size = new Size(380, 96);
			richTextBox1.TabIndex = 1;
			richTextBox1.Text = "";
			// 
			// Form1
			// 
			AutoScaleDimensions = new SizeF(7F, 15F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(800, 450);
			Controls.Add(richTextBox1);
			Controls.Add(DatabaseToExcel);
			Name = "Form1";
			Text = "Form1";
			ResumeLayout(false);
		}

		#endregion

		private Button DatabaseToExcel;
		private RichTextBox richTextBox1;
	}
}
