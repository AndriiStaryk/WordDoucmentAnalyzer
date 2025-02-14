using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;

public class DataGridViewCalendarColumn : DataGridViewColumn
{
    public DataGridViewCalendarColumn() : base(new DataGridViewCalendarCell())
    {
    }

    public override DataGridViewCell CellTemplate
    {
        get => base.CellTemplate;
        set
        {
            if (value is not DataGridViewCalendarCell)
            {
                throw new InvalidCastException("Must be a DataGridViewCalendarCell");
            }
            base.CellTemplate = value;
        }
    }
}

public class DataGridViewCalendarCell : DataGridViewTextBoxCell
{
    public DataGridViewCalendarCell() : base()
    {
        this.Style.Format = "d"; // Short date format
    }

    public override void InitializeEditingControl(int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
    {
        base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
        DataGridViewCalendarEditingControl ctl = DataGridView.EditingControl as DataGridViewCalendarEditingControl;

        if (this.Value is DateTime dateTimeValue)
        {
            ctl.Value = dateTimeValue;
        }
        else
        {
            ctl.Value = DateTime.Today;
        }
    }

    public override Type EditType => typeof(DataGridViewCalendarEditingControl);
    public override Type ValueType => typeof(DateTime);
    public override object DefaultNewRowValue => DateTime.Today;
}

public class DataGridViewCalendarEditingControl : DateTimePicker, IDataGridViewEditingControl
{
    private DataGridView dataGridView;
    private int rowIndex;
    private bool valueChanged = false;

    public DataGridViewCalendarEditingControl()
    {
        this.Format = DateTimePickerFormat.Short;
    }

    public object EditingControlFormattedValue
    {
        get => this.Value.ToShortDateString();
        set
        {
            if (DateTime.TryParse(value as string, out DateTime newDate))
            {
                this.Value = newDate;
            }
        }
    }

    public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context) => EditingControlFormattedValue;
    public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle) { }

    public int EditingControlRowIndex
    {
        get => rowIndex;
        set => rowIndex = value;
    }

    public bool EditingControlWantsInputKey(Keys key, bool dataGridViewWantsInputKey) => true;
    public void PrepareEditingControlForEdit(bool selectAll) { }
    public bool RepositionEditingControlOnValueChange => false;

    public DataGridView EditingControlDataGridView
    {
        get => dataGridView;
        set => dataGridView = value;
    }

    public bool EditingControlValueChanged
    {
        get => valueChanged;
        set => valueChanged = value;
    }

    public Cursor EditingPanelCursor => base.Cursor;

    protected override void OnValueChanged(EventArgs eventargs)
    {
        valueChanged = true;
        EditingControlDataGridView?.NotifyCurrentCellDirty(true);
        base.OnValueChanged(eventargs);
    }
}