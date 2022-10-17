using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using MySql.Data.MySqlClient;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace DBAdminTool
{
   
    partial class Form1
    {
       
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;


        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog3 = new System.Windows.Forms.OpenFileDialog();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.openFileDialog4 = new System.Windows.Forms.OpenFileDialog();
            this.label7 = new System.Windows.Forms.Label();
            this.button7 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(335, 81);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(67, 22);
            this.button1.TabIndex = 0;
            this.button1.Text = "Insert";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(335, 115);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(67, 22);
            this.button2.TabIndex = 1;
            this.button2.Text = "Export";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(119, 52);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(201, 20);
            this.dateTimePicker1.TabIndex = 2;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel Files New Format (*.xlsx)|*.xlsx|Excel Files Old Format (*.xls)|*.xls|All F" +
    "iles (*.*)|*.*";
            this.openFileDialog1.Title = "Select an excel file";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.Filter = "Image Files (*.jpg)|*.jpg|PNG Files (*.png)|*.png|BMP Files (*.bmp)|*.bmp|All Fil" +
    "es (*.*)|*.*";
            this.openFileDialog2.Title = "Select A Picture File";
            // 
            // openFileDialog3
            // 
            this.openFileDialog3.Filter = "Image Files (*.jpg)|*.jpg|PNG Files (*.png)|*.png|BMP Files (*.bmp)|*.bmp|All Fil" +
    "es (*.*)|*.*";
            this.openFileDialog3.Title = "Select A Picture File";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(119, 15);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(252, 20);
            this.textBox1.TabIndex = 3;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(377, 15);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(25, 20);
            this.button3.TabIndex = 6;
            this.button3.Text = "...";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 88);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(267, 15);
            this.label1.TabIndex = 9;
            this.label1.Text = "Press to insert the chosen data to the database:";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 122);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(233, 15);
            this.label2.TabIndex = 10;
            this.label2.Text = "Press to export a backup of the database:";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Source Excel File:";
            // 
            // openFileDialog4
            // 
            this.openFileDialog4.Filter = "Image Files (*.jpg)|*.jpg|PNG Files (*.png)|*.png|BMP Files (*.bmp)|*.bmp|All Fil" +
    "es (*.*)|*.*";
            this.openFileDialog4.Title = "Select A Picture File";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(13, 156);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(316, 15);
            this.label7.TabIndex = 18;
            this.label7.Text = "Press to Delete Data From Database for the Chosen Date:";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // button7
            // 
            this.button7.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button7.Location = new System.Drawing.Point(335, 149);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(67, 22);
            this.button7.TabIndex = 19;
            this.button7.Text = "Delete";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(416, 212);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.MaximumSize = new System.Drawing.Size(432, 250);
            this.MinimumSize = new System.Drawing.Size(432, 250);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Wind Database Admin Tool";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private Button button1;
        private Button button2;
        private DateTimePicker dateTimePicker1;
        private OpenFileDialog openFileDialog1;
        private OpenFileDialog openFileDialog2;
        private OpenFileDialog openFileDialog3;
        private TextBox textBox1;
        private Button button3;
        private Label label1;
        private Label label2;
        private Label label3;
        private OpenFileDialog openFileDialog4;
        private Label label7;
        private Button button7;
    }
}

