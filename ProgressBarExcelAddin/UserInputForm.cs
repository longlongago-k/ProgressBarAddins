using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProgressBarExcelAddin
{
  public partial class UserInputForm : Form
  {
    public float Value { get => ((ValueData)(bindingSourceValue.DataSource)).Value; }
    public UserInputForm()
    {
      InitializeComponent();
      bindingSourceValue.DataSource = new ValueData() { Value = 30 };
    }

    private void TrackBarValue_Scroll(object sender, EventArgs e)
    {

    }

    private void ButtonOK_Click(object sender, EventArgs e)
    {
      DialogResult = DialogResult.OK;
    }

    private void ButtonCancel_Click(object sender, EventArgs e)
    {
      DialogResult = DialogResult.Cancel;
    }

    private void UserInputForm_Load(object sender, EventArgs e)
    {
      try
      {
        textBoxValue.Focus();
      }catch(Exception ex)
      {
        System.Diagnostics.Debug.WriteLine(ex);

      }
    }

        private void textBoxValue_Validating(object sender, CancelEventArgs e)
        {
            if (float.TryParse(textBoxValue.Text, out float _) == false)
                e.Cancel = true;

        }

        private void bindingSourceValue_DataError(object sender, BindingManagerDataErrorEventArgs e)
        {

        }
    }

    public class ValueData : INotifyPropertyChanged
  {
    public event PropertyChangedEventHandler PropertyChanged;
    /// <summary>
    /// イベント通知
    /// </summary>
    /// <param name="info"></param>
    private void NotifyPropertyChanged(String info)
    {
      PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(info));
    }
    float _value;
    public float Value
    {
      get { return _value; }
      set
      {
        if (value != _value)
        {
          _value = value;
          // このプロパティ名を渡してイベント通知
          NotifyPropertyChanged("Value");
        }
      }
    }
  }
}
