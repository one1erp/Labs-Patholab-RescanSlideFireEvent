namespace RescanSlideFireEvent
{
    partial class RescanSlideFireEvent
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.listViewEntities = new System.Windows.Forms.ListView();
            this.close_button = new System.Windows.Forms.Button();
            this.Ok_button = new System.Windows.Forms.Button();
            this.txtEditEntity = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.listViewEntities);
            this.panel1.Controls.Add(this.close_button);
            this.panel1.Controls.Add(this.Ok_button);
            this.panel1.Controls.Add(this.txtEditEntity);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(13, 26);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(836, 459);
            this.panel1.TabIndex = 0;
            // 
            // listViewEntities
            // 
            this.listViewEntities.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.listViewEntities.GridLines = true;
            this.listViewEntities.Location = new System.Drawing.Point(64, 109);
            this.listViewEntities.Name = "listViewEntities";
            this.listViewEntities.Size = new System.Drawing.Size(731, 283);
            this.listViewEntities.TabIndex = 6;
            this.listViewEntities.UseCompatibleStateImageBehavior = false;
            this.listViewEntities.View = System.Windows.Forms.View.List;
            // 
            // close_button
            // 
            this.close_button.Location = new System.Drawing.Point(488, 404);
            this.close_button.Name = "close_button";
            this.close_button.Size = new System.Drawing.Size(75, 23);
            this.close_button.TabIndex = 5;
            this.close_button.Text = "Close";
            this.close_button.UseVisualStyleBackColor = true;
            this.close_button.Click += new System.EventHandler(this.close_button_Click);
            // 
            // Ok_button
            // 
            this.Ok_button.Location = new System.Drawing.Point(305, 404);
            this.Ok_button.Name = "Ok_button";
            this.Ok_button.Size = new System.Drawing.Size(75, 23);
            this.Ok_button.TabIndex = 4;
            this.Ok_button.Text = "OK";
            this.Ok_button.UseVisualStyleBackColor = true;
            this.Ok_button.Click += new System.EventHandler(this.Ok_button_Click);
            // 
            // txtEditEntity
            // 
            this.txtEditEntity.Location = new System.Drawing.Point(503, 70);
            this.txtEditEntity.Name = "txtEditEntity";
            this.txtEditEntity.Size = new System.Drawing.Size(137, 20);
            this.txtEditEntity.TabIndex = 3;
            this.txtEditEntity.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.aliquotName_KeyPress);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold);
            this.label2.Location = new System.Drawing.Point(647, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(156, 24);
            this.label2.TabIndex = 2;
            this.label2.Text = ":הזן מספר סלייד";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(355, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(209, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "סליידים לסריקה חוזרת";
            // 
            // RescanSlideFireEvent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Name = "RescanSlideFireEvent";
            this.Size = new System.Drawing.Size(871, 503);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtEditEntity;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListView listViewEntities;
        private System.Windows.Forms.Button close_button;
        private System.Windows.Forms.Button Ok_button;
    }
}
