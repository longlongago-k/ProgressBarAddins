namespace ProgressBarPowerPointAddin
{
  partial class UserInputForm
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
            this.components = new System.ComponentModel.Container();
            this.textBoxValue = new System.Windows.Forms.TextBox();
            this.bindingSourceValue = new System.Windows.Forms.BindingSource(this.components);
            this.trackBarValue = new System.Windows.Forms.TrackBar();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSourceValue)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.trackBarValue)).BeginInit();
            this.flowLayoutPanel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxValue
            // 
            this.textBoxValue.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.bindingSourceValue, "Value", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged, "0"));
            this.textBoxValue.Dock = System.Windows.Forms.DockStyle.Left;
            this.textBoxValue.Location = new System.Drawing.Point(0, 0);
            this.textBoxValue.Name = "textBoxValue";
            this.textBoxValue.Size = new System.Drawing.Size(47, 19);
            this.textBoxValue.TabIndex = 0;
            this.textBoxValue.Validating += new System.ComponentModel.CancelEventHandler(this.textBoxValue_Validating);
            // 
            // bindingSourceValue
            // 
            this.bindingSourceValue.DataSource = typeof(ProgressBarPowerPointAddin.ValueData);
            // 
            // trackBarValue
            // 
            this.trackBarValue.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSourceValue, "Value", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged, "0"));
            this.trackBarValue.Dock = System.Windows.Forms.DockStyle.Top;
            this.trackBarValue.LargeChange = 10;
            this.trackBarValue.Location = new System.Drawing.Point(47, 0);
            this.trackBarValue.Maximum = 100;
            this.trackBarValue.Name = "trackBarValue";
            this.trackBarValue.Size = new System.Drawing.Size(137, 45);
            this.trackBarValue.TabIndex = 1;
            this.trackBarValue.TabStop = false;
            this.trackBarValue.TickFrequency = 10;
            this.trackBarValue.Scroll += new System.EventHandler(this.TrackBarValue_Scroll);
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(106, 3);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.TabStop = false;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.ButtonCancel_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(25, 3);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 0;
            this.buttonOK.TabStop = false;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.ButtonOK_Click);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoSize = true;
            this.flowLayoutPanel1.Controls.Add(this.buttonCancel);
            this.flowLayoutPanel1.Controls.Add(this.buttonOK);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 60);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(184, 29);
            this.flowLayoutPanel1.TabIndex = 1;
            this.flowLayoutPanel1.WrapContents = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Top;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Progress(%)";
            // 
            // panel3
            // 
            this.panel3.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panel3.Controls.Add(this.trackBarValue);
            this.panel3.Controls.Add(this.textBoxValue);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 12);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(184, 45);
            this.panel3.TabIndex = 1;
            // 
            // UserInputForm
            // 
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(184, 89);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(200, 128);
            this.Name = "UserInputForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "進捗でFill";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.UserInputForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSourceValue)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.trackBarValue)).EndInit();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    #endregion
    private System.Windows.Forms.TextBox textBoxValue;
    private System.Windows.Forms.TrackBar trackBarValue;
    private System.Windows.Forms.BindingSource bindingSourceValue;
    private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
    private System.Windows.Forms.Button buttonCancel;
    private System.Windows.Forms.Button buttonOK;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Panel panel3;
    private System.Windows.Forms.ToolTip toolTip1;
  }
}