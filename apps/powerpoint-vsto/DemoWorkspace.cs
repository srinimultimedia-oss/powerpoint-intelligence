using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PowerPointIntelligence.AddIn
{
    internal sealed class DemoWorkspace : UserControl
    {
        private readonly ListBox list;
        private readonly Label description;
        private readonly TextBox search;
        private readonly ComboBox category;
        private readonly Button execute;
        private readonly Action<DemoTool> run;
        private readonly CheckBox inventory;
        internal readonly Label Status;
        internal DemoWorkspace(Action<DemoTool> run)
        {
            this.run=run; Dock=DockStyle.Fill; BackColor=Color.White; Font=new Font("Segoe UI",10); AutoScaleMode=AutoScaleMode.Font;
            var layout=new TableLayoutPanel{Dock=DockStyle.Fill,Padding=new Padding(12),ColumnCount=1,RowCount=9};
            foreach(var height in new[]{38,42,32,32,28})layout.RowStyles.Add(new RowStyle(SizeType.Absolute,height));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,96));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,40));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,64));
            layout.Controls.Add(new Label{Text="Build your presentation",Dock=DockStyle.Fill,Font=new Font("Segoe UI",14,FontStyle.Bold)},0,0);
            layout.Controls.Add(new Label{Text="Demonstration • sample data\nExternal services are not configured.",Dock=DockStyle.Fill},0,1);
            search=new TextBox{Dock=DockStyle.Fill,AccessibleName="Search tools and feature inventory"};layout.Controls.Add(search,0,2);
            category=new ComboBox{Dock=DockStyle.Fill,DropDownStyle=ComboBoxStyle.DropDownList,AccessibleName="Feature category"};layout.Controls.Add(category,0,3);
            inventory=new CheckBox{Text="Show full requested inventory",AutoSize=true};layout.Controls.Add(inventory,0,4);
            list=new ListBox{Dock=DockStyle.Fill,IntegralHeight=false,AccessibleName="Tools and features"};layout.Controls.Add(list,0,5);
            description=new Label{Dock=DockStyle.Fill,Padding=new Padding(0,8,0,0)};layout.Controls.Add(description,0,6);
            execute=new Button{Text="Run selected tool",Dock=DockStyle.Fill};layout.Controls.Add(execute,0,7);
            Status=new Label{Dock=DockStyle.Fill,Padding=new Padding(0,8,0,0),Text="Choose a tool. Changes use native editable objects."};layout.Controls.Add(Status,0,8);
            Controls.Add(layout);
            category.Items.Add("All categories");foreach(var item in DemoTools.Catalog.Select(t=>t.Category).Distinct().OrderBy(t=>t))category.Items.Add(item);category.SelectedIndex=0;
            search.TextChanged+=(s,e)=>RefreshList();category.SelectedIndexChanged+=(s,e)=>RefreshList();inventory.CheckedChanged+=(s,e)=>{category.SelectedIndex=0;RefreshList();};
            list.SelectedIndexChanged+=(s,e)=>Selection();execute.Click+=(s,e)=>{var tool=list.SelectedItem as DemoTool;if(tool!=null)this.run(tool);};
            var menu=new ContextMenuStrip();var runItem=menu.Items.Add("Run selected tool");runItem.Click+=(s,e)=>execute.PerformClick();menu.Opening+=(s,e)=>runItem.Enabled=execute.Enabled;list.ContextMenuStrip=menu;
            RefreshList();
        }
        private void RefreshList()
        {
            list.BeginUpdate();list.Items.Clear();string term=search.Text.Trim();
            if(inventory.Checked)
            {
                foreach(var feature in FeatureInventory.Items.Where(f=>(f.Name+" "+f.Category+" "+f.State).IndexOf(term,StringComparison.OrdinalIgnoreCase)>=0))list.Items.Add(feature);
            }
            else
            {
                foreach(var tool in DemoTools.Catalog.Where(t=>(category.SelectedIndex==0||t.Category==(string)category.SelectedItem)&&(t.Name+" "+t.Category).IndexOf(term,StringComparison.OrdinalIgnoreCase)>=0))list.Items.Add(tool);
            }
            list.EndUpdate();if(list.Items.Count>0)list.SelectedIndex=0;else Selection();category.Enabled=!inventory.Checked;
        }
        private void Selection()
        {
            var tool=list.SelectedItem as DemoTool;var feature=list.SelectedItem as FeatureInventoryEntry;
            description.Text=tool!=null?tool.Category+" • implemented demo, runtime unverified\n"+tool.Description:feature!=null?feature.Id+" • "+feature.State+"\n"+feature.Note:"No matching tools.";
            execute.Enabled=tool!=null;
        }
    }
}
