namespace RelatorioReport
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
            this.Enviar = new System.Windows.Forms.Button();
            this.Nome = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // Enviar
            // 
            this.Enviar.Location = new System.Drawing.Point(429, 51);
            this.Enviar.Name = "Enviar";
            this.Enviar.Size = new System.Drawing.Size(75, 23);
            this.Enviar.TabIndex = 0;
            this.Enviar.Text = "Enviar";
            this.Enviar.UseVisualStyleBackColor = true;
            this.Enviar.Click += new System.EventHandler(this.t_Click);
            // 
            // Nome
            // 
            this.Nome.Location = new System.Drawing.Point(59, 53);
            this.Nome.Name = "Nome";
            this.Nome.Size = new System.Drawing.Size(346, 20);
            this.Nome.TabIndex = 1;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(563, 261);
            this.Controls.Add(this.Nome);
            this.Controls.Add(this.Enviar);
            this.Name = "Form1";
            this.Text = "Formulario ";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Enviar;
        private System.Windows.Forms.TextBox Nome;
    }
}

